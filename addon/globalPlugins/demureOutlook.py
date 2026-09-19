# Copyright (C) 2026 TJOlsen
# SPDX-License-Identifier: GPL-2.0-or-later
# Demure Outlook for NVDA
# Announces metadata for newly delivered mail through Outlook's own event API.

import globalPluginHandler
import api
import core
import ui
import comHelper
import time
import os
import winsound
from collections import deque
import gui
import wx
import globalVars
from scriptHandler import script
from winAPI.sessionTracking import isLockScreenModeActive
from logHandler import log


class NotificationHistoryDialog(wx.Dialog):
	"""A stable snapshot so incoming mail never moves the selected notification."""

	def __init__(self, parent, messages):
		super().__init__(parent, title="Demure Outlook - Recent notifications", style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER)
		sizer = wx.BoxSizer(wx.VERTICAL)
		label = wx.StaticText(self, label="Recent &notifications, newest first. Use Up and Down to review; Escape to close.")
		sizer.Add(label, 0, wx.ALL, 10)
		self.notifications = wx.ListBox(self, choices=list(messages), style=wx.LB_SINGLE | wx.LB_HSCROLL)
		self.notifications.SetName("Recent notifications, newest first")
		sizer.Add(self.notifications, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 10)
		closeButton = wx.Button(self, wx.ID_CANCEL, label="&Close")
		sizer.Add(closeButton, 0, wx.ALIGN_RIGHT | wx.ALL, 10)
		self.SetSizer(sizer)
		self.SetSize((800, 300))
		self.SetMinSize((450, 200))
		self.SetEscapeId(wx.ID_CANCEL)
		self.CentreOnScreen()
		self.notifications.SetSelection(0)
		self.notifications.SetFocus()


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	"""Announce new mail without scanning the Inbox or reading message bodies."""

	def __init__(self):
		super().__init__()
		self._outlook = None
		self._connection = None
		self._connecting = False
		self._announcedEntryIDs = {}
		self._lastMailAt = None
		self._recentNotifications = deque(maxlen=5)
		self._historyDialog = None
		self._terminated = False
		self._connect()

	def terminate(self):
		self._terminated = True
		self._recentNotifications.clear()
		if self._historyDialog:
			self._historyDialog.EndModal(wx.ID_CANCEL)
		try:
			if self._connection:
				self._connection.disconnect()
		except Exception:
			pass
		self._connection = None
		self._outlook = None
		super().terminate()

	def _connect(self):
		if self._terminated or self._connection or self._connecting:
			return
		self._connecting = True
		try:
			from comtypes.client import GetEvents, GetModule
			outlook = self._getRunningOutlookObject()
			if not outlook:
				raise RuntimeError("Outlook's running object model is not ready")
			# Outlook exposes several event interfaces. NVDA's native object is
			# a dynamic COM object, so comtypes cannot infer one automatically.
			# ApplicationEvents_11 is the Outlook interface that carries NewMailEx.
			outlookTypes = GetModule(("{00062FFF-0000-0000-C000-000000000046}", 9, 6, 0))
			connection = GetEvents(outlook, self, interface=outlookTypes.ApplicationEvents_11)
			self._outlook = outlook
			self._connection = connection
			log.info("Demure Outlook: connected to Outlook's new-mail event")
		except Exception:
			log.warning("Demure Outlook: Outlook event connection is not ready; retrying", exc_info=True)
			core.callLater(15000, self._connect)
		finally:
			self._connecting = False

	@staticmethod
	def _getRunningOutlookObject():
		"""Attach to existing Outlook without waiting for an NVDA app module.

		NVDA's helper handles uiAccess privilege differences. GetActiveObject
		only attaches to a registered, running instance; it never starts Outlook.
		Avoid nativeOm's focus-juggling fallback so this stays in the background.
		"""
		return comHelper.getActiveObject("outlook.application", dynamic=True)

	def NewMailEx(self, this, entryIDCollection):
		"""Outlook event for mail that has just arrived; this is not Inbox polling."""
		if self._terminated:
			return
		now = time.monotonic()
		playSound = self._lastMailAt is None or now - self._lastMailAt >= 1200
		self._lastMailAt = now
		self._announcedEntryIDs = {
			entryID: announcedAt
			for entryID, announcedAt in self._announcedEntryIDs.items()
			if now - announcedAt < 120
		}
		for entryID in str(entryIDCollection).split(","):
			if entryID in self._announcedEntryIDs:
				log.info("Demure Outlook: ignored a duplicate new-mail event")
				continue
			try:
				item = self._outlook.Session.GetItemFromID(entryID)
				if playSound:
					self._playNewMailSound()
					playSound = False
				account = self._accountFor(item)
				sender = self._text(getattr(item, "SenderName", "")) or "Unknown sender"
				subject = self._text(getattr(item, "Subject", "")) or "No subject"
				message = "New mail for {0} from {1}. {2}.".format(account, sender, subject)
				core.callLater(0, self._announceAndRemember, message)
				self._announcedEntryIDs[entryID] = now
				log.info("Demure Outlook: announced a new message")
			except Exception:
				log.exception("Demure Outlook: could not announce a new message")

	def _announceAndRemember(self, message):
		if self._terminated:
			return
		self._recentNotifications.appendleft(message)
		if not globalVars.appArgs.secure and not isLockScreenModeActive():
			ui.message(message)

	@script(
		description="Show the last five Demure Outlook notifications, newest first",
		category="Demure Outlook",
		gesture="kb:control+alt+shift+o",
	)
	def script_recentNotifications(self, gesture):
		wx.CallAfter(self._showRecentNotifications)

	def _showRecentNotifications(self):
		if self._terminated or globalVars.appArgs.secure or isLockScreenModeActive():
			return
		if self._historyDialog:
			self._historyDialog.Raise()
			self._historyDialog.notifications.SetFocus()
			return
		if not self._recentNotifications:
			ui.message("No Demure Outlook notifications yet. History begins when NVDA starts.")
			return
		gui.mainFrame.prePopup()
		try:
			self._historyDialog = NotificationHistoryDialog(gui.mainFrame, tuple(self._recentNotifications))
			self._historyDialog.ShowModal()
		finally:
			if self._historyDialog:
				self._historyDialog.Destroy()
				self._historyDialog = None
			gui.mainFrame.postPopup()

	def _playNewMailSound(self):
		"""Play the familiar Windows email cue only while Outlook is in the background."""
		try:
			focus = api.getFocusObject()
			if focus and focus.appModule and focus.appModule.appName == "outlook":
				return
			windowsDirectory = os.environ.get("WINDIR", r"C:\\Windows")
			soundPath = os.path.join(windowsDirectory, "Media", "Windows Notify Email.wav")
			winsound.PlaySound(soundPath, winsound.SND_FILENAME | winsound.SND_ASYNC | winsound.SND_NODEFAULT)
		except Exception:
			log.debugWarning("Demure Outlook: could not play the new-mail sound", exc_info=True)

	def _accountFor(self, item):
		try:
			store = item.Parent.Store
			storeID = store.StoreID
			for account in self._outlook.Session.Accounts:
				try:
					if account.DeliveryStore.StoreID == storeID:
						return self._text(account.SmtpAddress) or self._text(account.DisplayName)
				except Exception:
					continue
			return self._text(store.DisplayName) or "Unknown account"
		except Exception:
			return "Unknown account"

	@staticmethod
	def _text(value):
		# Mail metadata is untrusted input. Remove control characters and cap it so
		# an unusually crafted message cannot produce an excessive announcement.
		cleaned = "".join(character if character.isprintable() else " " for character in str(value or ""))
		return " ".join(cleaned.split())[:300]
