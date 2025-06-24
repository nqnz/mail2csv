# mail2csv
Utilizes the microsoft graph API to read a user's mailbox and download attachments, timestamp them, and mark them as read. Originally aimed at CSV files. 

# usage 

export AZ_TENANT_ID="your-tenant-id"   
export AZ_CLIENT_ID="your-client-id"   
export AZ_CLIENT_SECRET="your-client-secret"   
export CSV_INGEST_EMAIL="csv-ingest@yourcompany.com"   
export DATA_DIR="path/to/data/csv"   
export STATE_FILE="path/to/data/.graph_delta.json"   
export R_SUBJECT="Daily Report" # Email Subject that we're searching on   

# install dependencies   

pip install msal requests  

# TODO
## docker/podman   
## Ask for variables at run time, ability to save those and run from a config file  
## Fix/verify the mark as read is working properly




