from pytube import YouTube
from youtube_search import YoutubeSearch
import xlrd
from os import path
# from pprint import pprint
from termcolor import colored, cprint
from moviepy.editor import *
import sys

youtube_url = 'https://youtube.com'
loc = "./2010s power hour playlist.xlsx"

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

beer_animation_title = 'Animated Hand(s) Opening Can of Beer ~ Green Screen'
into_video = ''
intro_length = 0
outro_video = ''
outro_length = 0
clip_length = 30
number_to_download = 20


def get_best_video(videos=[]):
    if len(videos) > 0:
        for video in videos:
            if video['title'].find('official'):
                return video
        return videos[0]
    else:
        return False


def video_downloaded(title: str, extension: str='mp4'):
    safe_title = make_tite_safe(title)
    c_title = colored(safe_title, 'red')
    print('trying to check if - ./' + c_title + '.' + extension + ' exists')
    return path.exists('./' + safe_title + '.' + extension)


def make_tite_safe(title: str):
    return title.replace('.', '').replace('"', '').replace('~', '').replace(',', '').replace('*', '')


def process_clip(clip, offset: int = 0):
    return clip.subclip(offset, offset + clip_length).resize((1280, 720))


def download_video(title: str, link):
    if not video_downloaded(title):
        cprint('DOWNLOADING VIDEO', 'red')
        YouTube(youtube_url + link).streams.get_highest_resolution().download()
    else:
        cprint('VIDEO already downloaded', 'yellow')
    return VideoFileClip('./' + make_tite_safe(title) + '.mp4')

# START
cprint('Checking BEER ANIMATION', 'green')
if not video_downloaded(beer_animation_title):
    cprint('Beer animation not found - ask your local editor to help you', 'red')
    YouTube('https://www.youtube.com/watch?v=z2TAKFJHJnA').streams.get_highest_resolution().download()
else:
    cprint('BEER ANIMATION already downloaded', 'yellow')
# make clip shorter
beer_title = make_tite_safe(beer_animation_title)
beer_animation_clip = VideoFileClip(beer_title + '.mp4')
beer_animation_clip = beer_animation_clip.subclip(3, 9).fx(vfx.mask_color, color=[0, 255, 0], thr=100, s=5)

count = 1
all_video_clips = []
all_beer_effects = []

while count < number_to_download + 1:
    cprint('SEARCHING FOR SONG "' + sheet.cell_value(count, 0) + '"', 'green')
    results = YoutubeSearch(sheet.cell_value(count, 0), max_results=5)
    video = get_best_video(results.videos)
    if video:
        print('Video found online: ' + video['title'] + video['link'])
        clip = download_video(video['title'], video['link'])
        short_clip = process_clip(clip, int(sheet.cell_value(count, 3)))

        print('Adding clip to all_video_clips')
        all_video_clips.append(short_clip)
        print('Adding beer animation')
        all_beer_effects.append(
            beer_animation_clip.set_start(intro_length + (clip_length * count) - 5)
            .set_position('center', 'center')
        )

    else:
        cprint('No good video found for ' + sheet.cell_value(count, 0), 'red')
        sys.exit()
    count += 1

cprint('Finished processing video - total of ' + str(len(all_video_clips)), 'green')
print('join all clips')
final_clip = concatenate_videoclips(all_video_clips)

print('Adding beer can masks')
masked_clip = CompositeVideoClip([
    final_clip.set_position('center', 'center'),
    *all_beer_effects,
], size=(1280, 720)).set_duration(intro_length + outro_length + (number_to_download * clip_length))

print('Writing file')
masked_clip.write_videofile('./final.mp4', fps=30)
