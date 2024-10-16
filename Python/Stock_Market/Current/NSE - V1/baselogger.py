import logging
import time

logger = logging
logger.basicConfig(filename='Nse_Data_'+time.strftime('%Y%m%d%H%M%S')+'.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')


