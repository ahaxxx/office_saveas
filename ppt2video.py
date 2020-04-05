# -*- coding: UTF-8 -*-
# -*- python:3.7    -*-
# -*- author:LiuBX  -*-

import win32com.client
import time
import os
import shutil

def pptx_to_mp4(ppt_path,mp4_target,resolution = 720,frames = 24,quality = 60,timeout = 120):
    # 状态初始化，且返回此值作为最终处理状态 0:失败  -1: 超时  1:成功
    status = 0
    if ppt_path == '' or mp4_target == '':
        return status
    # 处理开始时间
    start_tm = time.time()

    # 如果储存路径不存在则创建路径
    sdir = mp4_target[:mp4_target.rfind('\\')]
    if not os.path.exists(sdir):
        os.makedirs(sdir)
        
    ppt = win32com.client.Dispatch('PowerPoint.Application')
    presentation = ppt.Presentations.Open(ppt_path,WithWindow=False)
    presentation.CreateVideo(mp4_target,-1,1,resolution,frames,quality)
    while True:  
        try:
            time.sleep(0.1)
            if time.time() - start_tm > timeout:
                # 当处理超时时结束ppt进程且返回状态值-1
                os.system("taskkill /f /im POWERPNT.EXE")
                status = -1
                break
            if os.path.exists(mp4_target) and os.path.getsize(mp4_target) == 0:
                # 如果输出空文件则跳过本循环，返回状态0
                continue
            status = 1
            break
        except Exception as e:
            print('[错误]错误代码: {c}, 错误信息, {m}').format(c = type(e).__name__, m = str(e))
            break
    print(time.time()-start_tm)
    if status != -1:
        ppt.Quit()

    return status
    
def ppt2video_transform(ppt_path,video_path):

    # 画面质量
    quality = 60
    # 分辨率
    resolution = 720
    # 帧率
    frames = 24
    # ppt文件路径，只支持pptx
    ppt_path = os.path.abspath(ppt_path)
    # 转化视频储存路径
    mp4_path = os.path.abspath(video_path)

    ie_temp_dir = ''
    
    status = 0
    # 设置超时时间
    timeout = 4*60
    try:
        status = pptx_to_mp4(ppt_path,mp4_path,resolution,frames,quality,timeout)
        # 转换结束后清除缓存
        if ie_temp_dir != '':
            shutil.rmtree(ie_temp_dir, ignore_errors=True)
    except Exception as e:
        print('[错误]错误代码: {c}, 错误信息, {m}').format(c = type(e).__name__, m = str(e))
        
    if status == -1:
        print('[转换超时]转换失败')
    elif status == 1:
        print('转换成功!')
    else:
        if os.path.exists(mp4_path):
            os.remove(mp4_path)
        print('[失败]存在未知元素，请尝试手动转换')