import os
import inspect
import logging

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import xlwings as xw

from models import Artist, Album

this_dir = os.path.realpath(os.path.dirname(inspect.getfile(inspect.currentframe())))

# Logging
logging.basicConfig(filename=os.path.join(this_dir, 'xlwings-database.log'),
                    level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

log = logging.getLogger(__name__)

# SQLAlchemy boilerplate
engine = create_engine('sqlite:///{0}'.format(os.path.join(this_dir, 'chinook.sqlite')))
Session = sessionmaker(bind=engine)
session = Session()


def artist_query():
    xw.Workbook.caller()

    xw.Range('A4').table.clear_contents()
    query_string = '%{0}%'.format(xw.Range('B1').value)

    log.info('Performing query with: {0}'.format(query_string))

    artist_album = session.query(Artist.Name, Album.Title).join(Album).\
                           filter(Artist.Name.like(query_string))

    log.info('Query returned {0} records.'.format(artist_album.count()))

    try:
        xw.Range('A4').value = artist_album.all()
    except Exception as e:
        # E.g. will be triggered for queries that return 0 records
        log.exception(e)
        xw.Range('A4').value = 'An error occurred!'
