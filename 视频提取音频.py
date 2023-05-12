import office

# 这里填写你的视频位置
path = 'r' + input('需要提取音频的视频路径：')
# path，是你的视频位置；mp3_name，是你的MP3结果文件的名称，可以不填
office.video.video2mp3(path=path, mp3_name='result')