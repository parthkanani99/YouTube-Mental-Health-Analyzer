import os
import pandas as pd
from googleapiclient.discovery import build

# ==== CONFIGURATION ====
YOUTUBE_API_KEY = "AIzaSyAjIm9YAFzAaYUznIlGzCX0_G2kTk2B0xc"  # Replace with your actual YouTube API key
EXCEL_INPUT_FILE = r"C:\Users\parth\OneDrive\Documents\IDS 506 Mental health\.venv\YoutubeLinks.xlsx"
EXCEL_OUTPUT_FILE = r"C:\Users\parth\OneDrive\Documents\IDS 506 Mental health\.venv\YT_Extracted Data.xlsx"
EXCEL_INPUT_SHEET = "Sheet1"
EXCEL_OUTPUT_SHEET = "YouTubeData"
# =======================

def get_youtube_metadata(video_id):
    try:
        youtube = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)

        # Video details
        video_response = youtube.videos().list(part="snippet,statistics,contentDetails", id=video_id).execute()
        if not video_response["items"]:
            return ["Not found"] * 8

        video_data = video_response["items"][0]
        title = video_data["snippet"].get("title", "Not found")
        channel_id = video_data["snippet"].get("channelId", "Not found")
        channel = video_data["snippet"].get("channelTitle", "Not found")
        views = video_data["statistics"].get("viewCount", "Not found")
        likes = video_data["statistics"].get("likeCount", "Not found")
        duration = video_data["contentDetails"].get("duration", "Not found")
        comment_count = video_data["statistics"].get("commentCount", "Not found")

        # Channel details
        channel_response = youtube.channels().list(part="snippet,statistics", id=channel_id).execute()
        if channel_response["items"]:
            channel_data = channel_response["items"][0]
            channel_start_date = channel_data["snippet"].get("publishedAt", "Not found")
            subscriber_count = channel_data["statistics"].get("subscriberCount", "Not found")
        else:
            channel_start_date = "Not found"
            subscriber_count = "Not found"

        return [title, channel, views, likes, duration, subscriber_count, channel_start_date, comment_count]

    except Exception as e:
        print(f"❌ Error retrieving metadata for video {video_id}: {e}")
        return ["Error"] * 8

def load_video_urls_from_excel(file_path, sheet_name="Sheet1"):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return df.iloc[:, 0].dropna().tolist()

def extract_video_id(url):
    if "watch?v=" in url:
        return url.split("v=")[-1].split("&")[0]
    elif "youtu.be/" in url:
        return url.split("youtu.be/")[-1].split("?")[0]
    elif "embed/" in url:
        return url.split("embed/")[-1].split("?")[0]
    else:
        return None

def main():
    print("📥 Loading YouTube URLs from Excel...")
    urls = load_video_urls_from_excel(EXCEL_INPUT_FILE, EXCEL_INPUT_SHEET)
    data = []

    for url in urls:
        video_id = extract_video_id(url)
        if not video_id:
            print(f"⚠️ Skipping invalid URL: {url}")
            continue

        print(f"🔍 Fetching data for video ID: {video_id}")
        metadata = get_youtube_metadata(video_id)
        data.append([url] + metadata)

    # Define column headers
    base_columns = [
        "YouTube URL", "Title", "Channel", "Views", "Likes", "Video Duration",
        "Subscriber Count", "Channel Start Date", "Total Comments"
    ]

    df_output = pd.DataFrame(data, columns=base_columns)
    df_output.to_excel(EXCEL_OUTPUT_FILE, index=False, sheet_name=EXCEL_OUTPUT_SHEET)
    print(f"✅ Saved results to {EXCEL_OUTPUT_FILE}")

if __name__ == "__main__":
    main()
