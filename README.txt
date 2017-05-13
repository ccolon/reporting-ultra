Reporting Ultra
===============


# Setup the environment
- virtualenv env --no-site-packages
- . env/bin/activate
- pip install flask
- pip install xlsxwriter


# Run the application

```
# Activate the python environment
. env/bin/activate
# Configure the system environment
export FLASK_APP=server.py
export UPLOAD_FOLDER='/home/aperrin/lemontri'
# Run the server
flask run
```

# Local development
```
# Activate the python environment
. env/bin/activate
# Configure the system environment
export FLASK_APP=server.py
export UPLOAD_FOLDER=.
# Run the server
flask run
```



