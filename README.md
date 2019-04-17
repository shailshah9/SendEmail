# SendEmail
This is a simple python app that helps send emails to recruiters from the list of recruiters in google sheets. Also keeps track of previously send emails to avoid redundancy.

# Steps to follow:
1. Fill up the required variables in Constants.py file with your credentials
2. Replace my resume with yours and give proper name
3. Change the body of the email, just make sure to keep 2 %s in Body intact. 1st %s goes at Hello <Name of Recruiter> and 2nd %s goes for Company name. So draft your own body and just don't forget to put two %s in the body. :)

# Run:
1. pip3 install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
2. pip3 install oauth2client
3. python3 main.py

# Good Luck :)
