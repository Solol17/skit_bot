import logging
import os
import sys
from logging.handlers import RotatingFileHandler
import redis
from sys import platform

if "linux" in platform:
    geckodriver_path = '/action/workspace/src/geckodriver'
    firefox_location = "/usr/bin/firefox"
    redis_server = redis.Redis(host='redis_container', port=6379, db=0)
else:
    # Здесь прописать для Windows
    geckodriver_path = os.path.abspath("../geckodriver/geckodriver.exe")
    firefox_location = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"
    redis_server = redis.Redis(host='127.0.0.1', port=6379, db=0)

find_str = "С 1 по "
source_dir = os.path.abspath("sources")
document_path = os.path.abspath("skit_report.docx")


api_key = os.getenv("api_key")
model = "mistral-large-latest"

my_cache = dict()
log_dir = os.path.abspath("../Logs")
if not os.path.exists(log_dir):
    os.mkdir(log_dir)
logging.basicConfig(level=logging.INFO,  # Уровень логирования
                    handlers=[RotatingFileHandler(os.path.join(log_dir, "skit_bot_log.log"), maxBytes=5000000,
                                                  backupCount=10),
                              logging.StreamHandler(sys.stdout)],
                    format="%(asctime)s - [%(levelname)s] -  %(name)s - (%(filename)s).%(funcName)s(%(lineno)d) - %(message)s")
