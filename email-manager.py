import argparse
import json
import os
from datetime import datetime, timedelta
from O365 import Inbox

parser = argparse.ArgumentParser()


class EmailManager(object):
    def __init__(self, mail_folder):
        self.mail_folder = mail_folder
        self.username = 'qa.ex@office365.ecknhhk.xyz'
        self.password = 'ew68I7W52p*W'
        self.suffix = '.txt'

    def _get_twenty_four_hours_ago_datetime(self):
        """
        Generate today date in certain format
        :return: today date string
        """
        d = datetime.today() - timedelta(days=1)
        dt = d.strftime('%Y-%m-%dT%H:%M:%SZ')
        return dt

    def _format_email(self, email):
        """
        Format email to pretty JSON
        :param email: email object
        :return: pretty json
        """
        return json.dumps(email.json, indent=4, sort_keys=True)

    def _get_subject(self, email):
        """
        Get subject string
        :param email: email object
        :return: subject string
        """
        jjson = email.json
        subject = jjson.get('Subject')
        return str(subject.encode('utf-8'))

    def _get_filename(self, email):
        """
        Convert subject to filename
        :param email: email object
        :return: filename string
        """
        subject = self._get_subject(email)
        m_subject = subject.replace(' ', '_')
        m_subject = m_subject.replace('.', '')
        m_subject = m_subject.replace(':', '_')
        return m_subject.lower()

    def fetch_emails(self):
        """
        Fetches emails from O365 account
        :return: emails list
        """
        auth = (self.username, self.password)
        i = Inbox(auth, getNow=False)

        datetime_string = self._get_twenty_four_hours_ago_datetime()
        i.setFilter("DateTimeReceived gt {dt}".format(dt=datetime_string))
        i.getMessages(number=100)

        if len(i.messages) == 0:
            i.setFilter("")
            i.getMessages(number=10)
        return i.messages

    def write_file(self, filename, content):
        """
        Write email to file
        :param filename: filename
        :param content: file content
        """
        f = open(filename, "wb")
        f.write(content)
        f.close()

    def main(self):
        emails = self.fetch_emails()
        if not os.path.exists(self.mail_folder):
            os.makedirs(self.mail_folder)
        for email in emails:
            filename = self._get_filename(email)
            filepath = os.path.join(self.mail_folder, filename + self.suffix)
            content = self._format_email(email)
            self.write_file(filepath, content)

if __name__ == "__main__":
    parser.add_argument('--folder', dest='mail_folder', required=True, help='Fullpath to local destination folder')
    emailmanager = EmailManager(**vars(parser.parse_args()))
    emailmanager.main()
