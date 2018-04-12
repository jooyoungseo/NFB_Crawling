import csv
import mailbox
import html2text

# get body of email
def get_message(message):
    if message.is_multipart():
        body = ""
        for part in message.get_payload():
            if part.is_multipart():
                for subpart in part.walk():
                    if subpart.get_content_type() == "text/html":
                        body += str(subpart)
                    elif subpart.get_content_type() == "text/plain":
                        body += str(subpart)
            else:
                body += str(part)
        return body
    else:
        return message.get_payload()


export_file_name = "clean_mail.csv"
# create CSV file
with open(export_file_name, "wb") as csvFile:
    writer = csv.writer(csvFile)
    # create header row
    writer.writerow(["subject", "from", "date", "body"])
    for message in mailbox.mbox("C:/Users/JooYoung/test.mbox"):
        contents = get_message(message)
        contents = html2text.html2text(contents)
        writer.writerow([message["subject"], message["from"], message["date"], contents])

