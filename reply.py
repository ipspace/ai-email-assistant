#!/usr/bin/env python3
#

import os
import sys
import openai
import imaplib
import smtplib
import email
import yaml
import typing
import textwrap

def read_config(config_file: str) -> dict:
  try:
    with open(os.path.expanduser(config_file)) as stream:
      config = yaml.safe_load(stream)
      if not isinstance(config,dict):
        print(f'Cannot read YAML configuration from {config_file}')
        sys.exit(1)
  except Exception as ex:
    print(f'Cannot read configuration file {config_file}: {str(sys.exc_info()[1])}')
    sys.exit(1)

  return config

def get_ai_response(config: dict, text: str) -> str:
  openai.organization = config['openai']['org']
  openai.api_key      = config['openai']['key']

  text = f"{config['prompt']}\n======\n{text}"

  reply = openai.Completion.create(
    prompt=text,
    model='text-davinci-003',
    max_tokens=2000,
    temperature=1,
    presence_penalty=0.5,
    frequency_penalty=0.5)

  print(f'{type(reply)}\n{reply}')
  return(reply.choices[0].text)

def open_IMAP(config: dict) -> imaplib.IMAP4:
  mail = imaplib.IMAP4_SSL(host=config['email']['host'])
  mail.login(config['email']['user'],config['email']['password'])
  mail.select(config['email']['folder'])
  return mail

def fetch_email(mail: imaplib.IMAP4) -> dict:

  resp_code, mail_ids = mail.search(None, "ALL")
#  print(mail_ids)

  for mail_id in mail_ids[0].decode().split():
    print(f"Fetching message {mail_id} from the input folder")
    mail_data: typing.Any
    resp_code, mail_data = mail.fetch(mail_id, '(RFC822)')
    if not isinstance(mail_data,list):
      return { 'id': mail_id }
    msg = email.message_from_bytes(mail_data[0][1])
    m_from = msg.get('From')
    m_subj = msg.get('Subject')
    for part in msg.walk():
      if part.get_content_type() == "text/plain":
        m_text = part.as_string()

        if len(m_text) > 50:
          return { 'from': m_from, 'subject': m_subj, 'id': mail_id, 'body': m_text }

  return {}

  print(f'From: {m_from}\nSubject: {m_subj}\n\n{m_text}')

def delete_message(mail: imaplib.IMAP4, mail_id: str) -> None:
  resp_code, response = mail.store(mail_id, '+FLAGS', '\\Deleted')
  print(f'Deleted message ID {mail_id}: {response}')
  resp_code, response = mail.expunge()
  print(f'Expunged deleted messages: {response}')

def sendmail(config: dict, m_to: str, m_subj: str, m_body: str) -> None:
  mail = smtplib.SMTP_SSL(config['email']['smtp'])
  mail.login(config['email']['user'],config['email']['password'])

  msg = email.message.EmailMessage()
  msg.set_default_type("text/plain")
  msg["From"] = config['email']['from']
  msg["To"] = [ m_to ]
  msg["Bcc"] = config['email']['bcc']
  msg["Subject"] =  m_subj
  msg.set_content(m_body)

  response = mail.send_message(msg)
  print(f'Sending message to {m_to}: {response}')
  mail.quit()

def main() -> None:
  config = read_config('~/.email-assistant.yml')
  mail = open_IMAP(config)
  msg = fetch_email(mail)
  if not 'mail_id' in msg:
    return

  if not 'body' in msg:
    delete_message(mail,msg['mail_id'])
    return

  subject = msg['subject'].split('\n')[0]
  print(f"From: {msg['from']}\nSubj: {msg['subject']}\n\n{msg['body']}")

  reply = get_ai_response(config,msg['body'])
  body = f"{reply}\r\n\r\n{'=' * 60}\r\n{msg['body']}"

  print (f"{'=' * 80}\nReplying with...\n\n{body}")
  sendmail(config,msg['from'],f"Re: {subject}",body)
  delete_message(mail,msg['id'])
  mail.logout()

main()
