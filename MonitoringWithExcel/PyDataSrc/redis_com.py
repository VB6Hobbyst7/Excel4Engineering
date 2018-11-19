# -*- coding: utf-8 -*-

"""
  get data from Redis service through the COM object

"""

import pythoncom
import redis

try:
    conn = redis.Redis('localhost')
except:
    pass


class RedisRealtime:
    _public_methods_ = ["get_snapshot"]
    _reg_progid_ = "Redis.Snapshot"
    _reg_clsid_ = pythoncom.CreateGuid()

    def get_snapshot(self, tagid):
        htag = conn.hmget(tagid, 'value', 'ts')
        return float(htag[0].decode())

if __name__ == "__main__":
    # Run "python redisr_com.py"
    #   to register the COM server.
    # Run "python redis_com.py --unregister"
    #   to unregister it.
    print("Registering COM server...")
    import win32com.server.register
    win32com.server.register.UseCommandLine(RedisRealtime)
