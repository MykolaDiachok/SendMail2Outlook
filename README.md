# SendMail2Outlook

Этот crazy проект сделан для отправки зашифрованной и подписанной почты через outlook Основная проблема заключается в том, что рассылку делает программа запущенная под сервисом и по непонятнысм причинам Микрософт не сделала возможности работы с криптованием через COM. Или по крайней мере у меня не получилось реализовать....

Для работы заходим в outlook и в безопасности включаем шифорвание и подписание исходящей почты.
после этого отправляем в программу json файл 
SendMail2Outlook.exe -f="temp.json"


Структура файла:
 {
     "subject": "Mail subject",
     "body": "Text bode or HTML body mail",
     "To": [
         "mail1@site",
         "mailN@site"
           ],
     "Attachments": [
     {
         "filename": "NameOfFile.Type",
         "Base64": "Data file in Base64",
     }
     ]
 }


Программа возвращает состояния отправки почты:
     0 - Success
     2 - InvalidFilename
     3 - EncryptionProblems
     4 - OutlookRunProblem
     10 - UnknownError
     
     
