from datetime import datetime

import openpyxl as openpyxl
from pymongo import MongoClient
import re

wb = openpyxl.load_workbook("/Users/deepakranganathan/Downloads/User Migration - User Management.xlsx")

# sheets = wb.sheetnames
# print(wb.active)
# ws = wb[sheets[1]]
# print(ws)
# for n, sheet in enumerate(wb.worksheets):
#     print('Sheet Index:[{}], Title:{}'.format(n, sheet.title))

sheets = wb.sheetnames
sheet = wb[sheets[6]]
print(sheet)
data = []
for row in range(2, 10):
    # val = sheet.cell(row, 4).value
    print(sheet.cell(row,4).fill)
    print()
    print()


stageConnectionStr = "mongodb://monger:9HT3X99ZN2gk3ad@services-mongo.infra-dev.joveo.com:27000/admin?connectTimeoutMS=10000&authSource=admin&authMechanism=SCRAM-SHA-1&3t.uriVersion=3&3t.databases=admin&3t.alwaysShowAuthDB=true&3t.alwaysShowDBFromUserRole=true"
prodConnectionStr = "mongodb://readonly:872zU424C67TN2R@superprodmongo-0-pvt.infra.joveo.com:27017,superprodmongo-4-pvt.infra.joveo.com:27017,superprodmongo-5-pvt.infra.joveo.com:27017/admin?readPreference=secondary&connectTimeoutMS=10000&authSource=admin&authMechanism=SCRAM-SHA-1&3t.uriVersion=3&3t.connection.name=prod&3t.databases=admin&3t.alwaysShowAuthDB=true&3t.alwaysShowDBFromUserRole=true"

client1 = MongoClient(prodConnectionStr)
client2 = MongoClient(stageConnectionStr)

mycollection = client1.mojo.joveo_users

newcollection = client2.mojo.joveo_users_deepak_copy

# Restore documents from the source collection.
for a in mycollection.find(no_cursor_timeout=True):
    try:
        newcollection.insert(a)
        print(a)
    except:
        print('did not copy')


# for a in newcollection.find({"email":{ '$regex' : ".*hak.*"}}):
#     print(a)

# result = newcollection.find({"$and" : [{"scope": {"$size" : 1}},
#                                        {"email":{ '$regex' : ".*hak.*"}}]})
# print("new ")
# for a in result:
#     print(a['email'],end="    ")
#     print(a['scope'])

# result = newcollection.find({"scope": {"$size" : 1}})
#
# for a in result:
#     print(a['email'],end="    ")
#     print(a['scope'])

# //Use delete_one or delete_many instead
# result = newcollection.remove({"$and" : [{"scope": {"$size" : 1}},
#                                         {"email":{ '$regex' : ".*hak.*"}}]})


# newcollection.update({"email":{ '$regex' : ".*hak.*"}},
#                      {"$pull": { "scope": {"application" : "Mojo"}}})


# x = newcollection.find({ "$where": "this.scope.length > 2" })
# for a in x:
#     print(a)
#
#
# newcollection.update({"email":{ '$regex' : ".*hak.*"}})

sample = ["abc","def","ghi"]

newcollection.update({"email":{ '$regex' : ".*hak.*"}},
                     {"$push":{"scope":{ "id": "ter","application" : "Mojo","clients":[],"publishers":[] } }})