To run the application:
1. Create virtual environment - virtualenv venv
2. Activate virtual environment source venv/bin/activate
3. Install requirements - pip install -r requirements.txt
4. To store emails in local folder, run the application and provide path_to_folder- python email-manager.py --folder <path_to_folder>
5. To upload emails to S3 storage service, run the application with '--upload' flag