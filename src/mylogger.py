import datetime
import logging
import traceback


class MyLogger(object):
    path = './Error.log'
    logging.basicConfig(filename=path, level=logging.ERROR,
                        format='%(asctime)s %(levelname)s %(name)s %(message)s')

    @classmethod
    def wt(cls) -> None:
        # 获取错误信息
        error = traceback.format_exc()
        # 获取错误时间
        when = datetime.datetime.now()
        time_list = str(when).split(' ')
        time_list[1] = time_list[1].split('.')[0]
        time = time_list[0] + '@' + time_list[1] + '\n'
        logging.error(time + error)


LOGGER = MyLogger()
