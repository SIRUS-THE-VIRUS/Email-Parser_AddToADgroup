# Email-Parser_AddToADgroup


Email-Parser_AddToADgroup is a command line tool for automating the adding of users to an AD group after an approval email is sent for the request.

## Installation

Use git-clone to download and run with python3 

Required Libraries : 

[imapclient](https://pypi.org/project/IMAPClient/)

[mailparser](https://pypi.org/project/mail-parser/) 

[bs4](https://pypi.org/project/bs4/)

[re](https://pypi.org/project/re101/)

[pyad](https://pypi.org/project/pyad/)

[logging](https://docs.python.org/3/library/logging.html)

[time](https://docs.python.org/3/library/time.html)

[requests](https://pypi.org/project/requests/)

```bash
git clone https://github.com/SIRUS-THE-VIRUS/Email-Parser_AddToADgroup.git
```

## Usage

Parameters need to be filled into the code as needed. Run the code by going into the directory and using python3.

```bash
cd Email-Parser_AddToADgroup
python3 AD-Email-parser.py
```
## Note
The program will run forever until it loses connection to the mail server or AD instance more than 3 times

The program will listen to a specified mailbox for incoming requests to add a user to an AD group. 

The option is there to integrate with Power Automate to send an email response notifying that the user has been added to AD group.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://choosealicense.com/licenses/mit/)
