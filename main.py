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
            print(src+' has been successfully converted!')
            break
        except Exception:
            pass
    PowerPoint.Quit()
    return

if __name__ == "__main__":
    print("Please input the resolution you want for the videoes")
    resol = ""
    while resol not in [480, 720, 1080 ,2160]:
        resol = int(input("You can only choose from [480, 720, 1080, 2160] as the output resolution: "))
    ppt_srcs = files= os.listdir('ppt\\')
    start_time = time.time()
    for ppt_src in ppt_srcs:
        cov_ppt((os.getcwd()+'\\ppt\\'+ppt_src), os.getcwd()+'\\video\\'+ ppt_src[:-5] + '.mp4', resol)
    end_time = time.time()
    print('total cost: ' + str(end_time -start_time) + 's')