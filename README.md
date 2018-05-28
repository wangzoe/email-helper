# email-helper；
get email content and download them to excel；
主要尝试了两种不同的包来处理邮件：poplib和imap；
imap的功能更加强大一些，可以指定邮箱内的文件夹等；但是对于内容的读取，返回的内容更加复杂一点，需要更细致的解码，尤其在面对中文的时候；
