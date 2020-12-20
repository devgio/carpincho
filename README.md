# Carpincho
Carpincho is a tool to classify emails, process them and manipulate the attachments.

## Features
  - Classify emails based on filters
  - Download attachments to specific locations depending on the source
  - Manupulate the downloaded attachments (most often used to get the files ready to be ingested into a database)


## Requirements
* O365==2.0.11


## Use

* Load you configuration in the config.py file.

* Create your email classifiers and drop them in the email_classifiers folder. This files can be added and modified while the program is running.

* Run email_processing.py

```sh
$ python email_processing.py
```


### Todos
 - Add more email connectors. Right now it can only connect to O365, but this is just a matter of creating new connectors and dropping them in the email_connectors folder.
