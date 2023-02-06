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
import argparse

IMAP_session: typing.Any = None
DEBUG_FLAG: bool = False

def parse_CLI_args() -> argparse.Namespace:
  global DEBUG_FLAG
  parser = argparse.ArgumentParser(description='AI-driven email assistant')
  parser.add_argument('--config', dest='config', action='store', default='~/.email-assistant.yml',
                  help='configuration file')
  parser.add_argument('--test', dest='test', action='store',help='Test content')
  parser.add_argument('--debug', dest='debug', action='store_true',help=argparse.SUPPRESS)
  args = parser.parse_args()
  if args.debug:
    DEBUG_FLAG = True

  return args

def fail_miserably(t: str) -> None:
  print(f'{t}\n{str(sys.exc_info()[1])}')
  sys.exit(1)

def debug(t: str) -> None:
  global DEBUG_FLAG
  if DEBUG_FLAG:
    print(t)

def read_config(config_file: str) -> dict:
  try:
    with open(os.path.expanduser(config_file)) as stream:
      config = yaml.safe_load(stream)
      if not isinstance(config,dict):
        print(f'Cannot read YAML configuration from {config_file}')
        sys.exit(1)
  except:
    fail_miserably(f'Cannot read configuration file {config_file}')

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

  debug(f'{type(reply)}\n{reply}')
  return(reply.choices[0].text)

def open_IMAP(config: dict) -> imaplib.IMAP4:
  mail = imaplib.IMAP4_SSL(host=config['email']['host'])
  mail.login(config['email']['user'],config['email']['password'])
  mail.select(config['email']['folder'])
  return mail

def fetch_email(mail: imaplib.IMAP4) -> dict:

  resp_code, mail_ids = mail.search(None, "ALL")

  for mail_id in mail_ids[0].decode().split():
    debug(f"Fetching message {mail_id} from the input folder")
    mail_data: typing.Any
    resp_code, mail_data = mail.fetch(mail_id, '(RFC822)')
    if not isinstance(mail_data,list):
      return { 'id': mail_id }

    msg = email.message_from_bytes(mail_data[0][1])
    m_from = msg.get('From')
    m_subj = msg.get('Subject')
    for part in msg.walk():
      if part.get_content_type() == "text/plain":
        m_text = part.get_payload(decode=True).decode('utf-8')
        m_text = m_text.encode("ascii","ignore").decode('ascii')
        if len(m_text) > 50:
          return { 'from': m_from, 'subject': m_subj, 'id': mail_id, 'body': m_text }

    return { 'id': mail_id }

  return {}

def delete_message(mail: imaplib.IMAP4, mail_id: str) -> None:
  resp_code, response = mail.store(mail_id, '+FLAGS', '\\Deleted')
  debug(f'Deleted message ID {mail_id}: {response}')
  resp_code, response = mail.expunge()
  debug(f'Expunged deleted messages: {response}')

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
  print(f'Sent message to {m_to}: {response}')
  mail.quit()

def get_input_message(config: dict, args: argparse.Namespace) -> dict:
  global IMAP_session

  if args.test:
    try:
      with open(args.test,'r') as infile:
        return { 'body': infile.read() }
    except:
      fail_miserably(f'Cannot read test file {args.test}')

  try:
    IMAP_session = open_IMAP(config)
  except:
    fail_miserably('Cannot establish connection to IMAP server')

  try:
    msg = fetch_email(IMAP_session)
  except:
    fail_miserably('Cannot fetch email message')

  if not 'id' in msg:
    print(f'No messages to reply to, no fun today :(')
    return {}

  debug(str(msg))
  if not 'body' in msg:
    try:
      delete_message(IMAP_session,msg['mail_id'])
    except:
      fail_miserably(f"Cannot delete message# {msg['mail_id']}")
    return {}

  return msg

def main() -> None:
  args = parse_CLI_args()
  config = read_config(args.config)
  msg = get_input_message(config,args)
  if not msg:
    return

  try:
    reply = get_ai_response(config,msg['body'])
  except:
    fail_miserably('Cannot get a response from OpenAI server')

  if args.test:
    print(reply)
    return

  subject = msg['subject'].split('\n')[0]
  debug(f"From: {msg['from']}\nSubj: {msg['subject']}\n\n{msg['body']}")

  body = f"{reply}\r\n\r\n{'=' * 60}\r\n{msg['body']}"

  debug(f"{'=' * 80}\nReplying with...\n\n{body}")

  try:
    sendmail(config,msg['from'],f"Re: {subject}",body)
  except:
    fail_miserably(f"Cannot send the reply message to {msg['from']}")

  try:
    delete_message(IMAP_session,msg['id'])
    IMAP_session.logout()
  except:
    fail_miserably('Cannot delete message that I just replied to')

if __name__ == '__main__':
  main()
