import argparse
import json
import os
from datetime import datetime, timedelta
from O365 import Inbox

import boto
import boto.s3
from boto.s3.key import Key

AWS_ACCESS_KEY_ID = 'AKIAJXGFTPO37WU6APVA'
AWS_SECRET_ACCESS_KEY = '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'
bucket_name = 'interview-exercises'
region = 'eu-central-1'


parser = argparse.ArgumentParser()


class EmailManager(object):
    def __init__(self, mail_folder, upload, report):
        self.mail_folder = mail_folder
        self.upload = upload
        self.report = report
        self.username = 'qa.ex@office365.ecknhhk.xyz'
        self.password = 'ew68I7W52p*W'
        self.suffix = '.txt'
        if self.upload:
            self.tabledata = []
            self.conn = boto.s3.connect_to_region(region,
                                                  aws_access_key_id=AWS_ACCESS_KEY_ID,
                                                  aws_secret_access_key=AWS_SECRET_ACCESS_KEY,)

            bucket = self.conn.lookup(bucket_name)
            self.k = Key(bucket)
        else:
            if not os.path.exists(self.mail_folder):
                os.makedirs(self.mail_folder)

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

    def _get_time(self, email):
        """
        Get time string
        :param email: email object
        :return: time string
        """
        jjson = email.json
        time = jjson.get('DateTimeReceived')
        return str(time)

    def _get_from(self, email):
        """
        Get from string
        :param email: email object
        :return: from string
        """
        jjson = email.json
        from_dict = jjson.get('From')
        from_email = from_dict['EmailAddress'].get('Address')
        return str(from_email)

    def _get_to(self, email):
        """
        Get To string
        :param email: email object
        :return: To string
        """
        jjson = email.json
        to_dict = jjson.get('ToRecipients')
        to_email = to_dict[0]['EmailAddress'].get('Address')
        return str(to_email)

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
        print 'Downloading emails received during last 24 hours...'
        i.getMessages(number=100)
        print 'Downloaded {n} emails'.format(n=len(i.messages))

        if len(i.messages) == 0:
            i.setFilter("")
            print 'Downloading last 10 emails received...'
            i.getMessages(number=10)
            print 'Downloaded {n} emails'.format(n=len(i.messages))

        return i.messages

    def write_file(self, filename, content):
        """
        Write email to file
        :param filename: filename
        :param content: file content
        """
        print 'Saving file {f} on local storage'.format(f=filename)
        f = open(filename, "wb")
        f.write(content)
        f.close()
        print 'File {f} saved'.format(f=filename)

    def upload_file(self, filename, content):
        """
        Upload file to S3
        :param filename: filename
        :param content: file content
        :return: download_url string
        """
        print 'Uploading file {f} to S3...'.format(f=filename)
        self.k.key = filename
        self.k.set_contents_from_string(content)
        download_url = self.k.generate_url(expires_in=60)
        print 'File {f} uploaded to S3\nDownload URL: {url}\n'.format(f=filename, url=download_url)
        return download_url

    def generate_report(self, tabledata):
        """
        Generate HTML report
        :param tabledata: list of dicts, each dict represents a line
        """
        str_table = \
            "<html><table border='1'><tr><th>Time</th><th>From</th><th>To</th><th>Subject</th><th>S3 Link</th></tr>"
        for line in tabledata:
            str_rw = "<tr>" \
                     "<td>" + line.get('time') + "</td>" \
                     "<td>" + line.get('from') + "</td>" \
                     "<td>" + line.get('to') + "</td>" \
                     "<td>" + line.get('subject') + "</td>" \
                     "<td>" + line.get('link') + "</td>" \
                     "</tr>"
            str_table = str_table + str_rw
        str_table = str_table + "</table></html>"
        hs = open("report.html", 'w')
        hs.write(str_table)
        hs.close()

    def main(self):
        emails = self.fetch_emails()
        for email in emails:
            filename = self._get_filename(email)
            filepath = os.path.join(self.mail_folder, filename + self.suffix)
            content = self._format_email(email)
            if self.upload:
                download_url = self.upload_file(filename, content)
                email_metadata = {'time': self._get_time(email),
                                  'from': self._get_from(email),
                                  'to': self._get_to(email),
                                  'subject': self._get_subject(email),
                                  'link': str(download_url)}
                self.tabledata.append(email_metadata)
            else:
                self.write_file(filepath, content)
        if self.report:
            self.generate_report(self.tabledata)

if __name__ == "__main__":
    parser.add_argument('--folder', dest='mail_folder', required=True, help='Fullpath to local destination folder')
    parser.add_argument('--upload', dest='upload', action='store_true', help='Upload emails to S3 storage')
    parser.add_argument('--report', dest='report', action='store_true', help='Generate HTML report')
    emailmanager = EmailManager(**vars(parser.parse_args()))
    emailmanager.main()
