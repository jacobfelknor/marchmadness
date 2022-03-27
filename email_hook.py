import imaplib
import email

from fetch import fetch_data


def read_email_from_gmail():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login("***", "***")
    mail.select("inbox")

    result, data = mail.search(None, "(UNSEEN)")
    mail_ids = data[0]

    id_list = mail_ids.split()

    for email_id in id_list:
        result, data = mail.fetch(email_id.decode(), "(RFC822)")

        for response_part in data:
            if isinstance(response_part, tuple):
                # from_bytes, not from_string
                msg = email.message_from_bytes(response_part[1])
                email_subject = msg["subject"]
                email_from = msg["from"]
                print("From : " + email_from + "\n")
                print("Subject : " + email_subject + "\n")
                if "madness" in email_subject.lower():
                    return True

    return False


if read_email_from_gmail():
    # got a request for a new bracket
    # run fetch
    print("New excel request... ")
    fetch_data()
