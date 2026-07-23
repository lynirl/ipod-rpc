import json
import os
import tempfile
import time

import win32com.client
import pywintypes

import requests
from dotenv import load_dotenv

from PIL import Image
from pypresence import Presence, exceptions as pypresence_exceptions
from pypresence.types import ActivityType

# ---- config ----------------------------------------------------------

#getting the data from our trusted dotenv™
load_dotenv()
CLIENT_ID = os.environ.get("DISCORD_CLIENT_ID")
CATBOX_USERHASH = os.environ.get("CATBOX_USERHASH", "").strip()

#get the pictures from the dev portal
SOURCE_IMAGES = {
    "Aqua": "aqua_ipod",
    "Obsidian": "pauline_ipod",
    "Library": "itunes_library",
}
#self-explanatory
DEFAULT_SOURCE_IMAGE = "itunes_library"
NO_ARTWORK_IMAGE = "no_image"

#toggle for uploading artwork
#if false; will use placeholder image all the time
UPLOAD_ARTWORK = True
CATBOX_API_URL = "https://catbox.moe/user/api.php"
#cache file will be located at temp
ARTWORK_CACHE_FILE = os.path.join(tempfile.gettempdir(), "artwork_cache.json")

#interval of refreshing
POLL_INTERVAL_SECONDS = 5
# ------------------------------------------------------------------------

#enums for the states
IT_PLAYER_STATE_STOPPED = 0
IT_PLAYER_STATE_PLAYING = 1
#the sources
IT_SOURCE_KIND_LIBRARY = 1
IT_SOURCE_KIND_IPOD = 2
IT_SOURCE_KIND_AUDIO_CD = 3
IT_SOURCE_KIND_MP3_CD = 4
#and all the pic stuff, what format we wanna export it to
IT_ARTWORK_FORMAT_BMP = 1
IT_ARTWORK_FORMAT_JPEG = 2
IT_ARTWORK_FORMAT_PNG = 3
ARTWORK_EXT = {IT_ARTWORK_FORMAT_BMP: "bmp", IT_ARTWORK_FORMAT_JPEG: "jpg", IT_ARTWORK_FORMAT_PNG: "png"}

#load image cache
def load_cache():
    try:
        with open(ARTWORK_CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}

#save image cache
def save_cache(cache):
    try:
        with open(ARTWORK_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except OSError:
        pass

#we're gripping that itunes tight
def get_itunes():
    try:
        return win32com.client.Dispatch("iTunes.Application")
    except pywintypes.com_error:
        return None

#by default, our source is the Library, unless we can find it's from an iPod
def read_source_info(itunes):
    source_name = "Library"
    try:
        source = itunes.CurrentPlaylist.Source
        if source.Kind == IT_SOURCE_KIND_IPOD:
            source_name = source.Name
    except (pywintypes.com_error, AttributeError):
        pass
    return source_name

#all the stuff to upload to catbox
#debug included for now; until i stop having issues with it
def upload_to_catbox(filepath):
    data = {"reqtype": "fileupload"}
    data["userhash"] = CATBOX_USERHASH

    file_size = os.path.getsize(filepath)
    print(f"  uploading {filepath} ({file_size} bytes)...")

    with open(filepath, "rb") as f:
        resp = requests.post(
            CATBOX_API_URL,
            data=data,
            files={"fileToUpload": f},
            timeout=20,
        )

    print(f"  catbox status={resp.status_code} content-type={resp.headers.get('content-type')} "
          f"body={resp.text!r}")

    resp.raise_for_status()
    url = resp.text.strip()

    #if we don't get an url back, that's an issue
    if not url.startswith("http"):
        raise ValueError(f"unexpected catbox response: {url!r}")
    return url

#returns the file url of the uploaded artwork
def get_artwork_url(track, cache):
    if not UPLOAD_ARTWORK:
        return None

    #if we don't know the artist/album then placeholder
    try:
        artist = track.Artist or "Unknown"
        album = track.Album or "Unknown"
        cache_key = f"{artist}::{album}"
    except pywintypes.com_error:
        return None

    #get cache status
    if cache_key in cache:
        cached = cache[cache_key]
        return None if cached == "NONE" else cached

    try:
        #get that art
        artwork_collection = track.Artwork
        if artwork_collection.Count == 0:
            cache[cache_key] = "NONE"
            save_cache(cache)
            return None
        artwork = artwork_collection.Item(1)
        fmt = artwork.Format
        ext = ARTWORK_EXT.get(fmt, "jpg")

    #don't cache if we have a problem
    except (pywintypes.com_error, AttributeError):
        return None 

    #raw (itunes stuff) and png (converted)
    raw_path = os.path.join(tempfile.gettempdir(), f"itunes_rpc_art_raw.{ext}")
    png_path = os.path.join(tempfile.gettempdir(), "itunes_rpc_art.png")
    url = None
    try:
        artwork.SaveArtworkToFile(raw_path)

        #since itunes gets weird with pictures, we use pillow to convert them
        with Image.open(raw_path) as img:
            img.load()
            print(f"  decoded artwork: {img.size} {img.mode}, source format {ext}")
            img.convert("RGB").save(png_path, "PNG")
        url = upload_to_catbox(png_path)

    #if it fails we keep the files for analysis, it uses the "no artwork" pic
    #it will try again later though if it falls on the same song
    except (pywintypes.com_error, requests.RequestException, ValueError, OSError) as e:
        print(f"artwork upload failed, icon fallback: {e}")
        print(f"  files kept for inspection: {raw_path} , {png_path}")
        cache[cache_key] = "NONE"
        save_cache(cache)
        return None

    #and well if it doesn't fail then we delete those files
    for p in (raw_path, png_path):
        try:
            os.remove(p)
        except OSError:
            pass

    cache[cache_key] = url
    save_cache(cache)
    return url

#where's the track at?
def read_track_state(itunes, cache):
    try:
        state = itunes.PlayerState
    except pywintypes.com_error:
        return None

    #if we ain't playing then we're either paused or stopped
    #but the effect is the same
    if state != IT_PLAYER_STATE_PLAYING:
        return None

    #get all the good stuff
    try:
        track = itunes.CurrentTrack
        name = track.Name
        artist = track.Artist
        album = track.Album
        duration = int(track.Duration)
        position = int(itunes.PlayerPosition)
    except pywintypes.com_error:
        return None

    source_name = read_source_info(itunes)
    artwork_url = get_artwork_url(track, cache)

    #all the info we need to build our presence
    return {
        "name": name or "Unknown Track",
        "artist": artist or "Unknown Artist",
        "album": album or "",
        "duration": duration,
        "position": position,
        "playing": state == IT_PLAYER_STATE_PLAYING,
        "source": source_name,
        "artwork_url": artwork_url,
    }

#build the actual presence!!
def build_presence_payload(track):
    now = time.time()
    #get all of our pictures; for the badge + for the song pic
    source_icon = SOURCE_IMAGES.get(track["source"], DEFAULT_SOURCE_IMAGE)
    big_image = track["artwork_url"] or NO_ARTWORK_IMAGE

    #what we'll be sending to the presence
    #everything happens here
    payload = {
        "details": track["name"],
        "state": f"by {track['artist']}" if track["artist"] else None,
        "large_image": big_image,
        "small_image": source_icon,
        "small_text": f"Listening from {track['source']}",
        "large_text": f"{track['album']}" if track["album"] else None,
        "activity_type": ActivityType.LISTENING,
    }

    if track["playing"] and track["duration"] > 0:
        start = now - track["position"]
        end = start + track["duration"]
        payload["start"] = int(start)
        payload["end"] = int(end)

    return payload

#track state changes
#"if old is none and new is none"
def state_changed(old, new):
    if old is None and new is None:
        return False
    if (old is None) != (new is None):
        return True
    return (
        old["name"] != new["name"]
        or old["artist"] != new["artist"]
        or old["playing"] != new["playing"]
        or old["source"] != new["source"]
    )

#the main. what can i say
def main():
    rpc = Presence(CLIENT_ID)
    rpc.connect()
    print("connected to Discord!!")

    cache = load_cache()
    last_track = None

    try:
        while True:
            itunes = get_itunes()
            if itunes is None:
                if last_track is not None:
                    rpc.clear()
                    last_track = None
                print("iTunes not running, waiting")
                time.sleep(POLL_INTERVAL_SECONDS)
                continue

            current = read_track_state(itunes, cache)

            if state_changed(last_track, current):
                if current is None:
                    rpc.clear()
                    print("no music playing, stopping presence")
                else:
                    #build the presence
                    payload = build_presence_payload(current)
                    #send it to discord
                    rpc.update(**{k: v for k, v in payload.items() if v is not None})
                    status = "Playing"
                    #"art" is if we have a picture; "icon" is if we're using the placefolder
                    art = "art" if current["artwork_url"] else "icon"
                    #shows every action in terminal
                    print(f"{status}: {current['name']} - {current['artist']} [{current['source']}, {art}]")

            #switching tracks
            last_track = current
            #we sleep for the set interval, before refreshing
            time.sleep(POLL_INTERVAL_SECONDS)

    #on ctrl-c
    except KeyboardInterrupt:
        print("bye bye!!! clearing everything")
        rpc.clear()
        rpc.close()
    except pypresence_exceptions.PipeClosed:
        print("no access to Discord (check if it's running)")


if __name__ == "__main__":
    main()