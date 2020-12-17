import win32com.client
import time
import os

def cov_ppt(src, dst, resol):
    PowerPoint = win32com.client.Dispatch('PowerPoint.Application')
    target=PowerPoint.Presentations.Open(src,WithWindow=False)
    target.CreateVideo(dst, VertResolution=resol)
    while True:
        time.sleep(1)
        try:
            os.rename(dst,dst)
            print(src+' 已经转换成功啦！')
            break
        except Exception:
            pass
    PowerPoint.Quit()
    return

if __name__ == "__main__":
    print("请输入你想要的输出视频分辨率")
    resol = ""
    while resol not in [480, 720, 1080 ,2160]:
        resol = int(input("你只能从 [480, 720, 1080, 2160] 这几个分辨率中选哦: "))
    ppt_srcs = files= os.listdir('ppt\\')
    print("开始转换啦!")
    start_time = time.time()
    for ppt_src in ppt_srcs:
        cov_ppt((os.getcwd()+'\\ppt\\'+ppt_src), os.getcwd()+'\\video\\'+ ppt_src[:-5] + '.mp4', resol)
    end_time = time.time()
    print('整个转换过程花了: ' + str(end_time -start_time) + 's呢')
    input("按任意键退出……")