# SQL-Query
I wrote this to query a SQL server and output the result into a .csv file. This saves a few steps and makes copying and pasting easier.

Both files were made to automate two separate queries to two separate SQL servers. Running these scripts saves time and eliminates the need for opening SQL and copying the query.
Both write the data obtained from the query to a CSV file.

The newest version of this script ('UHCsqlQuerytoCSV with GrantAccess.py') prompts for a username and password, uses Selenium to automate clicks on a webpage to grant my account read access to the SQL database, queries that database to find a list of jobs running for the day, and writes the output of that query to a CSV file.
