from common.utils.db_tools import db

sql1 = """
        select * from companyinfo
"""
db.session.execute(sql1)