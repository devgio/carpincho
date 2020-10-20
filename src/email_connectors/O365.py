import warnings
with warnings.catch_warnings():
	warnings.filterwarnings("ignore", category=DeprecationWarning)
	import os, shutil, O365
	dir_path = os.path.dirname(os.path.realpath(__file__))
	try:
		from .logger import *
	except:
		import logging

class Mailbox():
	'''
	This class connects to an Azure server App that then uses to read the emails. For mor info see documentation on O365 module -> Authentication.

	The azure_app_credentials variable is unique to the Azure App.

	It has to be initialized with a list in the format of [folder', 'subfolder', ...]. This indicates the email subfolder to get the emails from.
	'''
	def __init__(self, folder, azure_app_credentials, shared_mailbox=False):
		if shared_mailbox:
			self.account = O365.Account(azure_app_credentials, main_resource=shared_mailbox, timeout=60)
		else:
			self.account = O365.Account(azure_app_credentials, timeout=60)

		if not self.account.is_authenticated:  # will check if there is a token and has not expired
			self.account.authenticate(scopes=['basic', 'message_all', 'message_all_shared'])

		self.mailbox = self.get_folder(folder)

	def get_folder(self, folder):
		# Goes through the 'folder' list to define the mailbox were it will be reading messages from.
		mailbox = self.account.mailbox()
		for child in folder:
			mailbox = mailbox.get_folder(folder_name=child)
		return mailbox
	
	def messages(self, limit=200, batch=50):
		return self.mailbox.get_messages(limit=limit, batch=batch)

	def saveAttachments(self, message, save_files_to, file_condition=False, name_file_using=False, modify_file_using=False):
		if not message.has_attachments:
			return False
		if 'unknown' in str(message.attachments):
			message.attachments.download_attachments()

		for attachment in message.attachments:
			temp_file_path = os.path.join(dir_path, attachment.name)
			if not attachment.save(location=dir_path):
				continue

			if file_condition:
				if not file_condition(temp_file_path):
					os.remove(temp_file_path)
					continue

			if name_file_using:
				file_name = name_file_using(temp_file_path)
			else:
				file_name = attachment.name

			file_path = os.path.join(save_files_to, file_name)
			try:
				shutil.move(temp_file_path, file_path)
			except:
				os.remove(temp_file_path)
				return

			if modify_file_using: modify_file_using(file_path)


	def sendEmail(self, to, subject='', body='', cc=[], attachments=[]):
		try:
			# to must be an iterable of email addresses.
			m = self.account.new_message()
			for receipient in to:
				m.to.add(receipient)
			for receipient in cc:
				m.cc.add(receipient)
			m.subject = subject
			m.body = body
			for attachment in attachments:
				m.attachments.add(attachment)
			return m.send()
		except Exception as e:
			logging.warning(e)
			
	def forward(self, message, to):
		m = message.forward()
		for recipient in to:
			m.to.add(recipient)
		m.send()
