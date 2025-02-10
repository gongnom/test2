import pymysql
import sys, os
# sys.path.append(os.path.dirname(os.path.abspath(os.path.dirname(__file__))))   #부모경로를 참조 path에 추가
# from Config import config             # Config/config.py 임포트
import logging

class MyDB:
    """Database connection class."""

    def __init__(self, config):
        self.host = config['host']
        self.username = config['user']
        self.password = config['password']
        self.port = config['port']
        self.dbname = config['dbname']
        self.conn = None 

        print("\n현재 디비는 {} : {}입니다.\n".format(self.host, self.dbname))
    
    def open_connection(self):
        """Connect to MySQL Database."""
        try:
            if self.conn is None:
                self.conn = pymysql.connect(
                    host=self.host,
                    port=self.port,
                    user=self.username,
                    passwd=self.password,
                    db=self.dbname,
                    charset='utf8',
                    connect_timeout=5
                    )
        except pymysql.MySQLError as e:
            logging.error(e)
            sys.exit()
        finally:
            logging.info('Connection opened successfully.')

    def run_query(self, query, data=''):
        """Execute SQL query."""
        try:
            self.open_connection()
            with self.conn.cursor() as cur:
                if 'SELECT' in query:
                    records = []
                    if data:
                        cur.execute(query, data)
                    else:    
                        cur.execute(query)
                    result = cur.fetchall()
                    for row in result:
                        records.append(row)
                    cur.close()
                    return records
                else:
                    if data:
                        result = cur.execute(query, data)
                    else:
                        result = cur.execute(query)
                    self.conn.commit()
                    # affected = f"{cur.rowcount} rows affected."
                    cur.close()
                    return cur.rowcount
        except pymysql.MySQLError as e:
            print(e)
        finally:
            if self.conn:
                self.conn.close()
                self.conn = None
                logging.info('Database connection closed.')