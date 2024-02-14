import win32com.client


def create_outlook_draft(subject, bodies, signature, recipients, attachment_paths=None, image_paths=None):
    outlook = win32com.client.Dispatch("outlook.application")
    mail = outlook.CreateItem(0x0)
    mail.Subject = subject
    mail.To = recipients
    mail.BodyFormat = 2  # 2 means HTML format
    body1 = bodies[0]
    body2 = bodies[1]
    body3 = bodies[2]

    html_text1 = f"<p>{body1}</p>"
    html_text2 = f"<p>{body2}</p>"
    html_text3 = f"<p>{body3}</p>"
    html_img1 = f"<img src=\"cid:image1\">"
    html_img2 = f"<img src=\"cid:image2\">"
    html_img3 = f"<img src=\"cid:image3\">"
    signature = signature.replace("\n","<br>")
    html_signature = f"<p>{signature}</p>"
    mail.HTMLBody = (html_text1 + html_img1 + html_text2 + html_img2 + html_text3 + html_img3 + html_signature)

    if attachment_paths:
        for i in range(len(attachment_paths)):
            try:
                mail.Attachments.Add(attachment_paths[i])
            except:
                print("Error attach file")

    if image_paths:
        # Add image attachment
        for i in range(len(image_paths)):
            try:
                attachment = mail.Attachments.Add(image_paths[i])
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E",
                                                    "image" + str(i + 1))
            except:
                print("Error attach image")

    mail.Display()

# Example usage
if __name__ == "__main__":
    subject = 'Test Email with Attachment'

    name = "My name is Hanmeimei"
    age = "??"
    body1 = f"My name is {name}, <br>I'm {age} years old"
    body2 = f"Ahh {name}, I'm beautiful"
    body3 = f"Ohh {name}"
    bodies = [body1, body2, body3]
    signature = f"Best regards\n{name}\nSr..GSM"
    recipients = 'xxxxxxx@gmail.com'
    attachment_path1 = "C:/Users/xxx/Downloads/test_doc.txt"
    attachment_path2 = "C:/Users/xxx/Downloads/test_doc.txt"
    attachment_path3 = "C:/Users/xxx/Downloads/test_doc.txt"
    attachment_paths = [attachment_path1, attachment_path2, attachment_path3]
    image_path1 = "C:/Users/xxx/OneDrive/图片/屏幕快照/Screenshot 2023-09-28 153700.png"
    image_path2 = "C:/Users/xxx/OneDrive/图片/屏幕快照/Screenshot 2023-09-28 153700.png"
    image_path3 = "C:/Users/xxx/OneDrive/图片/屏幕快照/Screenshot 2023-09-28 153700.png"
    image_paths = [image_path1, image_path2, image_path3]

    create_outlook_draft(subject, bodies, signature, recipients, attachment_paths, image_paths)
