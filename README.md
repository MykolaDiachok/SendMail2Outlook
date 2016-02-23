<p># <strong>SendMail2Outlook</strong></p>
<p>Этот crazy проект сделан для отправки зашифрованной и подписанной почты через outlook Основная проблема заключается в том, что рассылку делает программа запущенная под сервисом и по непонятным причинам Микрософт не сделала возможности работы с криптованием через COM. Или по крайней мере у меня не получилось реализовать....</p>
<p>Для работы заходим в outlook и в безопасности включаем шифорвание и подписание исходящей почты.<br />после этого отправляем в программу json файл <br />SendMail2Outlook.exe -f="temp.json"</p>
<p>Структура файла:</p>
<p>{<br />&nbsp;&nbsp;&nbsp;"subject": "Mail subject",<br />&nbsp;&nbsp;&nbsp;"body": "Text bode or HTML body mail",<br />&nbsp;&nbsp;&nbsp;"To": [<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"mail1@site",<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"mailN@site"<br />&nbsp;&nbsp;&nbsp;],<br />&nbsp;&nbsp;&nbsp;"Attachments": [<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"filename": "NameOfFile.Type",<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Base64": "Data file in Base64"<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;}<br />&nbsp;&nbsp;&nbsp;]<br />}</p>
<p><br />Программа возвращает состояния отправки почты:<br /> 0 - Success<br /> 2 - InvalidFilename<br /> 3 - EncryptionProblems<br /> 4 - OutlookRunProblem<br /> 10 - UnknownError<br /> </p>

<p>This project is made crazy to send encrypted and signed mail through outlook The main problem lies in the fact that the e-mail program does launched a service and for unknown reasons, Microsoft does not make it possible to work with encryption through COM. Or at least I did not get to realize ....</p>

To go to work in the security outlook and include encrypted and signing outgoing mail.
then send the program json file
<br />SendMail2Outlook.exe -f = "temp.json"</p>
<p>Struct of file:</p>
<p>{<br />&nbsp;&nbsp;&nbsp;"subject": "Mail subject",<br />&nbsp;&nbsp;&nbsp;"body": "Text bode or HTML body mail",<br />&nbsp;&nbsp;&nbsp;"To": [<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"mail1@site",<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"mailN@site"<br />&nbsp;&nbsp;&nbsp;],<br />&nbsp;&nbsp;&nbsp;"Attachments": [<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"filename": "NameOfFile.Type",<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"Base64": "Data file in Base64"<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;}<br />&nbsp;&nbsp;&nbsp;]<br />}</p>
<p><br />Return value:<br /> 0 - Success<br /> 2 - InvalidFilename<br /> 3 - EncryptionProblems<br /> 4 - OutlookRunProblem<br /> 10 - UnknownError<br /> </p>
