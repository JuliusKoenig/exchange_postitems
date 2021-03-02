# exchange_postitems
This script allows you to publish PostItems to the public folders of a Microsoft Exchange Server.

### mode selection:
```
usage: exchange_postitems.py [-h] [-V] [-v] {mode_auto,mode_manual} ...

positional arguments:
  {mode_auto,mode_manual}
    mode_auto                                   Automatic config driven mode.
    mode_manual                                 Manual mode.

optional arguments:
  -h, --help                                    show this help message and exit
  -V, --version                                 show program's version number and exit
  -v, --verbose                                 Show debugging information.
```
### mode_auto:
```
usage: exchange_postitems.py mode_auto [-h] -c CONFIG_FILE [--author AUTHOR] [--subject SUBJECT]

required arguments:
  -c CONFIG_FILE, --config_file CONFIG_FILE     A file with all the necessary configuration
                                                parameters. If it does not exist, it will be created
                                                during the next run.

optional arguments:
  -h, --help                                    show this help message and exit
  --author AUTHOR                               the author in new post
  --subject SUBJECT                             the subject in new post
```
### mode_manual:
```
usage: exchange_postitems.py mode_manual [-h] --server SERVER --username USERNAME --password PASSWORD 
                                              --primary_smtp_address PRIMARY_SMTP_ADDRESS 
                                              --target_path TARGET_PATH --author AUTHOR --subject SUBJECT

required arguments:
  --server SERVER                               MS Exchangeserver
  --username USERNAME                           the logon username
  --password PASSWORD                           the logon password
  --primary_smtp_address PRIMARY_SMTP_ADDRESS   the primary smtp address for this account.
  --ssl_verify                                  If this parameter is set, SSL verification is skipped.
  --target_path TARGET_PATH                     the target path in public folders
  --author AUTHOR                               the author in new post
  --subject SUBJECT                             the subject in new post

optional arguments:
  -h, --help                                    show this help message and exit
```

### Sample Configuration File:
```
{
  "server": "127.0.0.1",
  "username": "example.local\\user1",
  "password": "mysecret",
  "primary_smtp_address": "user1@example.local",
  "ssl_verify": true,
  "target_path": "path/to/public/folder",
  "author": "user1",
  "subject": "Hey it works!"
}
```
