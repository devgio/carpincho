from .logger import *
from .mailbox import mailbox
from config import download_folder
import os, importlib

# Functions for importing the classifiers
def _load_module(path):
	name = os.path.split(path)[-1]
	spec = importlib.util.spec_from_file_location(name, path)
	module = importlib.util.module_from_spec(spec)
	spec.loader.exec_module(module)
	return module

def import_classifiers_from(ECfolder):
	EC_list = []
	for fname in os.listdir(ECfolder):
		if not fname.startswith('.') and fname.endswith('.py'):
			try:
				module = _load_module(os.path.join(ECfolder, fname))
				EC_list.append(module.EC)
			except Exception as e:
				logging.error(f'import_classifiers_from ERROR {fname} failed to load:\n{e}')
				pass
	EC_list.sort(reverse=True, key=lambda ec: ec.priority_level)
	return EC_list


# Default file conditions
def isEXCEL(file):
	return file.lower().endswith(('xlsx', 'xls', 'xlsm'))

def isCSV(file):
	return file.lower().endswith('.csv')

def isXML(file):
	return file.lower().endswith('.xml')

default_cond_functions = {
	'isEXCEL':isEXCEL,
	'isCSV':isCSV,
	'isXML':isXML,
}


# Classifier class to be imported in each classifier file.
class EmailClassifier():
	def __init__(self, 
		move_email_to, 
		subject_contains=('',), 
		body_contains=('',), 
		filename_contains=(),
		addressed_to=(), 
		allowed_senders=('',), 
		file_condition=None, 
		name_file_using=None, 
		modify_file_using=None,
		save_files_to='', 
		what_to_do=None, 
		forward_to=None,
		priority_level=-10):

		# Filters
		self.allowed_senders = [x.lower() for x in allowed_senders]
		self.subject_contains = [x.lower() for x in subject_contains]
		self.body_contains = [x.lower() for x in body_contains]
		self.filename_contains = [x.lower() for x in filename_contains]
		self.addressed_to = [x.lower() for x in addressed_to]

		# Variables and functions for attachment download
		self.save_files_to = save_files_to
		if not os.path.isdir(self.save_files_to):
			self.save_files_to = os.path.join(download_folder, self.save_files_to)
		self.file_condition = file_condition
		if type(self.file_condition) == str:
			self.file_condition = default_cond_functions[self.file_condition]
		self.name_file_using = name_file_using
		self.modify_file_using = modify_file_using

		# Variables for the email handling
		self.move_email_to = move_email_to
		self.what_to_do = what_to_do
		self.forward_to = forward_to
		self.priority_level = priority_level

	def recognizes(self, message):
		# Compare the message against all the filters
		if not True in [x in str(message.subject).lower() for x in self.subject_contains]:
			return False

		if not True in [x in str(message.body).lower() for x in self.body_contains]:
			return False

		if not True in [sender in str(message.sender).lower() for sender in self.allowed_senders]:
			return False

		if self.filename_contains:
			if 'unknown' in str(message.attachments):
				message.attachments.download_attachments()
			if not True in [posible_name in str(attachment.name).lower() for posible_name in self.filename_contains for attachment in message.attachments]:
				return False

		if self.addressed_to:
			addressed_emails = [x.address.lower() for x in message.to] + [x.address.lower() for x in message.cc]
			if not True in [address in addressed_emails for address in self.addressed_to]:
				return False

		return True

	def _download_attachments(self, message):
		try:
			mailbox.saveAttachments(
				message=message, 
				save_files_to=self.save_files_to,
				file_condition=self.file_condition,
				name_file_using=self.name_file_using,
				modify_file_using=self.modify_file_using)
			return True
		except Exception as e:
			logging.warning(e)
			return False

	def action(self, message):
		if self.forward_to:
			mailbox.forward(message, self.forward_to)
		if self.what_to_do:
			try: 
				self.what_to_do(self, message)
				return True
			except Exception as e: 
				logging.error(f'what_to_do FAILED!\n- message: {message.subject}\n- exception: {e}')
				return False
		if self.save_files_to:
			return self._download_attachments(message)
		return True	


