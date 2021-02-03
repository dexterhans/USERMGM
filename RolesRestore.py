from pymongo import MongoClient

stageConnectionStr = "mongodb://mms-automation:GNYbrmR99XkJHPtAEmxSFyCr@superprodmongo-0-pvt.infra.joveo.com:27017,superprodmongo-4-pvt.infra.joveo.com:27017,superprodmongo-5-pvt.infra.joveo.com:27017/admin?readPreference=secondary&connectTimeoutMS=10000&authSource=admin&authMechanism=SCRAM-SHA-1"

client2 = MongoClient(stageConnectionStr)

database = client2.mojo
database.drop_collection(database.roles)


collection = database.roles_backup
collection.aggregate([{ '$match': {} }, {'$out': "roles"}])