import asyncio
import multiprocessing
from threading import Thread
from types import FunctionType


class AsyncRun(object):

    def __init__(self):
        self.loop = asyncio.new_event_loop()

    def start_loop(self):
        asyncio.set_event_loop(self.loop)
        try:
            self.loop.run_forever()
        finally:
            self.loop.run_until_complete(self.loop.shutdown_asyncgens())
            self.loop.close()

    def end_loop(self):
        self.loop.stop()
        self.loop.close()

    # 单进程io密集型
    def message_run_start(self, name=""):
        t = Thread(target=self.start_loop, name=name)
        t.start()

    # TODO: 多进程计算密集型，不完善
    @staticmethod
    def multiprocess_run(processes_num=1, target=None, args=()):
        pool = multiprocessing.Pool(processes=processes_num)
        results = []
        if target is None:
            print("请输入信息")
            return
        for i in range(processes_num):
            if type(target is FunctionType):
                result = pool.apply_async(target, args)  #直接运行函数
            results.append(result)
        pool.close()
        pool.join()
        # for i in results:
        #     if i.ready():  # 进程函数是否已经启动了
        #         if i.successful():  # 进程函数是否执行成功
        #             print(i.get())  # 进程函数返回值


# sample
def fib(n):
    index = 0
    a = 0
    b = 1
    while index < n:
        a, b = b, a + b
        index += 1
    return str(b)


async def async_fib(n):
    index = 0
    a = 0
    b = 1
    while index < n:
        asyncio.sleep(0.01)
        a, b = b, a + b
        index += 1
    print(b)


if __name__ == '__main__':
    # sample
    run = AsyncRun()
    # 多进程计算密集型
    run.multiprocess_run(3, fib, (2,))
    # 单进程io密集型
    run.message_run_start("io密集型")
    asyncio.run_coroutine_threadsafe(async_fib(3), run.loop)


