# coding:utf-8
import configparser
import os

# 用os模块来读取
curpath = os.path.dirname(os.path.realpath(__file__))
cfgpath = os.path.join(curpath, 'data.ini')  # 读取到本机的配置文件
dat = configparser.ConfigParser()
dat.read(cfgpath)
sender = dat.get('email_wx', 'sender')
print(sender)

if dat.has_section('emali_tel') is False:
    dat.add_section('emali_tel')
print(dat.sections())
dat.set('emali_tel', 'sender', 'yoyo1@tel.com')
dat.set('emali_tel', 'port', '265')
items = dat.items('emali_tel')
print(items)  # list里面对象是元祖
dat.write(open(cfgpath, 'w'))
