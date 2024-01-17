import requests
import pandas as pd
import urllib.request
from PIL import Image

# *************************************************************
token = "xxxxxxxxxx"
course = 0
# *************************************************************

# read poll csv
# df = pd.read_csv("ed_thread_image_poll.csv").fillna("")
df = pd.read_excel("ed_thread_image_poll.xlsx", engine="openpyxl", sheet_name="threads").fillna("")

for index, row in df.iterrows(): 
    endpoint_polls = f"https://eu.edstem.org/api/courses/{course}/polls"
    endpoint_threads = f"https://eu.edstem.org/api/courses/{course}/threads"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json; charset=utf-8"}

    #thread
    thread_category = row["thread_category"]
    thread_subcategory = row["thread_subcategory"]
    thread_subsubcategory = row["thread_subsubcategory"]
    thread_title = row["thread_title"]
    thread_body_1 = row["thread_body_1"]

    # image
    image = row["image_url"]
    if image != "": 
        urllib.request.urlretrieve(image, "image")  
        img = Image.open("image")
        sizeX, sizeY = img.size
    else:
        sizeX, sizeY = "", ""

    #poll
    poll_question = row["poll_question"]
    # poll options
    for index, row in df.iterrows(): 
        poll_options = [i for i in [row["poll_option_1"], row["poll_option_2"], row["poll_option_3"], row["poll_option_4"], row["poll_option_5"], row["poll_option_6"], row["poll_option_7"], row["poll_option_8"], row["poll_option_9"], row["poll_option_10"], row["poll_option_11"], row["poll_option_12"]] if i != ""]
    # poll settings
    multiple_choice = row["multiple_choice"]
    reveal_votes = row["reveal_votes"]
    pinned = row["pinned"]
    private = row["private"]
    megathread = row["megathread"]

    poll_options_formatted = [{"content":f"<document version=\"1.0\"><paragraph>{option}</paragraph></document>"} for option in poll_options]
    poll = {"content":f"<document version=\"1.0\"><paragraph>{poll_question}</paragraph></document>","multiple_choice":multiple_choice,"reveal_votes":reveal_votes,"closes_at":None,"options":poll_options_formatted}
    poll_id = requests.post(endpoint_polls, json=poll, headers=headers).json()["poll"]["id"]

    #thread (text after poll)
    thread_body_2 = row["thread_body_2"]

    #thread with image and poll
    thread = {
        "thread": {
            "type": "post",
            "title": f"{thread_title}",
            "category": f"{thread_category}",
            "subcategory": f"{thread_subcategory}",
            "subsubcategory": f"{thread_subsubcategory}",
            "content": f"<document version=\"2.0\"><paragraph>{thread_body_1}</paragraph><figure><image src=\"{image}\" width=\"{sizeX}\" height=\"{sizeY}\"/></figure><poll id=\"{poll_id}\"/><paragraph>{thread_body_2}</paragraph></document>",
            "is_pinned": pinned,
            "is_private": private,
            "is_anonymous": False,
            "is_megathread": megathread,
            "anonymous_comments": False
        }
    }

    # send info to ed
    requests.post(endpoint_threads, json=thread, headers=headers).json()

