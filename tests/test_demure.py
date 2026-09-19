import importlib.util
import sys
import types
import unittest
from unittest.mock import Mock
from pathlib import Path

S = types.SimpleNamespace
queue = []
def later(delay, fn, *args):
    queue.append((delay, fn, args))
def flush():
    while queue:
        delay, fn, args = queue.pop(0)
        if delay == 0:
            fn(*args)
class Base:
    def __init__(self): pass
    def terminate(self): pass
speech = Mock()
frame = S(prePopup=Mock(), postPopup=Mock())
wx = types.ModuleType('wx')
wx.Dialog = object
wx.ID_CANCEL = 5101
wx.CallAfter = lambda fn: later(0, fn)
def script(**kwargs):
    def decorate(fn):
        fn.metadata = kwargs
        return fn
    return decorate
modules = {
    'globalPluginHandler': S(GlobalPlugin=Base),
    'api': S(getFocusObject=lambda: S(appModule=S(appName='outlook'))),
    'core': S(callLater=later), 'ui': S(message=speech),
    'appModuleHandler': S(runningTable={}),
    'comHelper': S(getActiveObject=Mock()),
    'gui': S(mainFrame=frame), 'wx': wx,
    'globalVars': S(appArgs=S(secure=False)),
    'scriptHandler': S(script=script),
    'winAPI': types.ModuleType('winAPI'),
    'winAPI.sessionTracking': S(isLockScreenModeActive=lambda: False),
    'logHandler': S(log=Mock()),
    'winsound': S(PlaySound=Mock(), SND_FILENAME=1, SND_ASYNC=2, SND_NODEFAULT=4),
}
sys.modules.update(modules)
spec = importlib.util.spec_from_file_location('demure', Path(__file__).resolve().parents[1] / 'addon/globalPlugins/demureOutlook.py')
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
class Tests(unittest.TestCase):
    def setUp(self):
        queue.clear(); speech.reset_mock(); frame.prePopup.reset_mock(); frame.postPopup.reset_mock()
        m.isLockScreenModeActive = lambda: False
        m.globalVars.appArgs.secure = False
        m.comHelper.getActiveObject.reset_mock(side_effect=True, return_value=True)
        self.originalConnect = m.GlobalPlugin._connect
        m.GlobalPlugin._connect = lambda self: None
        self.p = m.GlobalPlugin()
        m.GlobalPlugin._connect = self.originalConnect
        store = S(StoreID='store', DisplayName='Test account')
        self.items = {str(i): S(Parent=S(Store=store), SenderName='Sender', Subject='Subject '+str(i)) for i in range(8)}
        self.getItem = Mock(side_effect=lambda i:self.items[i])
        self.p._outlook = S(Session=S(GetItemFromID=self.getItem, Accounts=[]))
    def test_bound_shortcut(self):
        self.assertEqual(m.GlobalPlugin.script_recentNotifications.metadata['gesture'], 'kb:control+alt+shift+o')
    def test_background_connection_without_app_module(self):
        outlook = object()
        connection = Mock()
        m.comHelper.getActiveObject.return_value = outlook
        events = Mock(return_value=connection)
        sys.modules['comtypes'] = types.ModuleType('comtypes')
        sys.modules['comtypes.client'] = S(GetEvents=events, GetModule=Mock(return_value=S(ApplicationEvents_11='events')))
        self.p._connect()
        m.comHelper.getActiveObject.assert_called_once_with('outlook.application', dynamic=True)
        self.assertIs(self.p._outlook, outlook)
        self.assertIs(self.p._connection, connection)
        self.assertFalse(queue)
        self.p._connect()
        self.assertEqual(events.call_count, 1)
    def test_outlook_not_running_retries_then_connects(self):
        events = Mock(return_value=Mock())
        sys.modules['comtypes'] = types.ModuleType('comtypes')
        sys.modules['comtypes.client'] = S(GetEvents=events, GetModule=Mock(return_value=S(ApplicationEvents_11='events')))
        m.comHelper.getActiveObject.side_effect = RuntimeError('Not running')
        self.p._connect()
        self.assertIsNone(self.p._connection)
        self.assertFalse(self.p._connecting)
        self.assertEqual(len(queue),1)
        delay, callback, args = queue.pop()
        self.assertEqual(delay,15000)
        events.assert_not_called()
        m.comHelper.getActiveObject.side_effect=None
        callback(*args)
        self.assertIsNotNone(self.p._connection)
    def test_bounded_newest_first_and_duplicate(self):
        for i in range(7): self.p.NewMailEx(None, str(i))
        self.p.NewMailEx(None, '6'); flush()
        self.assertEqual(len(self.p._recentNotifications), 5)
        self.assertEqual([x.split('Subject ')[1] for x in self.p._recentNotifications], ['6.','5.','4.','3.','2.'])
        self.assertEqual(speech.call_count, 7)
        self.assertEqual(self.getItem.call_count, 7)
    def test_failed_event_not_saved(self):
        self.p.NewMailEx(None,'missing'); flush()
        self.assertFalse(self.p._recentNotifications)
        speech.assert_not_called()
    def test_metadata_sanitized(self):
        self.items['0'].Subject = 'hello\r\n' + 'x'*400
        self.p.NewMailEx(None,'0'); flush()
        message = self.p._recentNotifications[0]
        self.assertNotIn('\n', message)
        self.assertEqual(message, speech.call_args.args[0])
        self.assertLess(len(message), 650)
    def test_empty_history(self):
        self.p.script_recentNotifications(None); flush()
        self.assertIn('No Demure',speech.call_args.args[0])
        frame.prePopup.assert_not_called()
    def test_snapshot_and_dialog_cleanup(self):
        self.p._announceAndRemember('Older'); self.p._announceAndRemember('Newer')
        dialogs=[]
        p=self.p
        class Dialog:
            def __init__(self,parent,messages):
                self.messages=messages; self.destroyed=False; self.notifications=S(SetFocus=Mock()); dialogs.append(self)
            def ShowModal(self): p._announceAndRemember('Arrived while reviewing')
            def Destroy(self): self.destroyed=True
        original=m.NotificationHistoryDialog; m.NotificationHistoryDialog=Dialog
        try: self.p._showRecentNotifications()
        finally: m.NotificationHistoryDialog=original
        self.assertEqual(dialogs[0].messages,('Newer','Older'))
        self.assertEqual(self.p._recentNotifications[0],'Arrived while reviewing')
        self.assertTrue(dialogs[0].destroyed); self.assertIsNone(self.p._historyDialog)
        frame.prePopup.assert_called_once(); frame.postPopup.assert_called_once()
    def test_existing_dialog_reused(self):
        dialog=S(Raise=Mock(),notifications=S(SetFocus=Mock()))
        self.p._historyDialog=dialog; self.p._showRecentNotifications()
        dialog.Raise.assert_called_once(); frame.prePopup.assert_not_called()
    def test_shutdown_ignores_queued_delivery_and_retry(self):
        self.p.NewMailEx(None,'0'); self.p.terminate(); flush()
        self.assertFalse(self.p._recentNotifications); speech.assert_not_called()
        self.p._connect(); self.assertFalse(queue)
    def test_lock_and_secure_mode(self):
        m.isLockScreenModeActive=lambda:True
        self.p._announceAndRemember('Private mail'); self.p._showRecentNotifications()
        speech.assert_not_called(); frame.prePopup.assert_not_called()
        m.isLockScreenModeActive=lambda:False; m.globalVars.appArgs.secure=True
        self.p._showRecentNotifications(); frame.prePopup.assert_not_called()
    def test_restart_empty(self):
        self.p._announceAndRemember('Mail'); self.p.terminate()
        self.assertFalse(self.p._recentNotifications)
if __name__ == '__main__': unittest.main(verbosity=2)
