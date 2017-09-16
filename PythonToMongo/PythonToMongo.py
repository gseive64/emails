
# first program with MongoDB
from pymongo import MongoClient
import pprint
client = MongoClient()
db = client.myEmails
emails = db.myEmailsTest5
pprint.pprint(emails.find_one())
#for email in emails.find():
#    pprint.pprint(email)
pipeline = [
    {"$project": {"_id": 0, "unread": 1, "date":1, "year": {"$year": "$date"},  "week": {"$week": "$date"} }},	
    {"$group": {"_id": {"week":"$week", "unread": "$unread"}, "count": {"$sum": 1}}}
    ]
read_unread_count = list(emails.aggregate(pipeline))

pprint.pprint(read_unread_count)