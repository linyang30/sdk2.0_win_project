import unittest
import logging
import logging.config
from common.common_func import get_config_data, mkdir
import shutil

mkdir('../log')
CON_LOG = '../config/log.conf'
logging.config.fileConfig(CON_LOG)
logging = logging.getLogger()

mkdir('../log')


class MyUnit(unittest.TestCase):

    data = get_config_data()

    def setUp(self):
        logging.info('=' * 20 + 'start test' + '=' * 20)

    def tearDown(self):
        logging.info('=' * 20 + 'end test' + '=' * 20)
        shutil.rmtree('../temp')