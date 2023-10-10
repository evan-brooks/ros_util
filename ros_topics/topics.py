import asyncio
import re
import subprocess
import webbrowser

from sheets import *

DEFAULT_SAMPLE_COUNT = 10
NO_RESPONSE_TIMEOUT = 15

TIMEOUT_SENTINEL = object()

def get_topic_list() -> list:
    command = ["ros2", "topic", "list"]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    return result.stdout.strip().split('\n')

async def run_command(command):
    process = await asyncio.create_subprocess_exec(
        *command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT
    )
    while True:
        try:
            line = await asyncio.wait_for(process.stdout.readline(), NO_RESPONSE_TIMEOUT)
            if not line:  # EOF
                break
            yield line.decode('utf-8')
        except asyncio.TimeoutError:
            yield TIMEOUT_SENTINEL

async def get_average_bandwidth(topic: str, samples: int = DEFAULT_SAMPLE_COUNT) -> float:

    command = ["ros2", "topic", "bw", topic]
    bandwidth_pattern = re.compile(r"(\d+\.\d+|\d+) (\wB?)/s")

    total_bandwidth = 0
    sample_count = 0

    first = True

    async for line in run_command(command):
        if line is TIMEOUT_SENTINEL:
            break

        match = bandwidth_pattern.search(line)
        if match:
            if first:
                first = False
                continue
            value, unit = match.groups()
            value = float(value)
            if unit == "GB":
                value = value * 8589934592 / (1024 * 1024)  
            elif unit == "MB":
                value = value * 8388608 / (1024 * 1024) 
            elif unit == "KB":
                value = value * 8192 / (1024 * 1024)
            elif unit == "B":
                value = value * 8 / (1024 * 1024)
            total_bandwidth += value
            sample_count += 1

        if sample_count >= samples:
            break

    average_bandwidth = total_bandwidth / samples if sample_count > 0 else 0
    return round(average_bandwidth, 3)


async def get_average_rate(topic: str, samples: int = DEFAULT_SAMPLE_COUNT) -> float:
    command = ["ros2", "topic", "hz", topic]
    rate_pattern = re.compile(r"average rate: (\d+\.\d+)")

    total_rate = 0
    sample_count = 0

    first = True

    async for line in run_command(command):
        if line is TIMEOUT_SENTINEL:
            break

        match = rate_pattern.search(line)
        if match:
            if first:
                first = False
                continue

            rate = float(match.group(1))
            total_rate += rate
            sample_count += 1

        if sample_count >= samples:
            break

    average_rate = total_rate / samples if sample_count > 0 else 0
    return round(average_rate, 1)

def get_topic_type(topic: str) -> str:
    command = ["ros2", "topic", "type", topic]
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    
    return result.stdout.strip()


async def main():
    topics = get_topic_list()
    topic_data = [] 
    print(f"collecting data for {len(topics)} topics . . .\n")

    print(f"{'TOPIC':<50}{'TYPE':<50}{'mbps':>8}{'hz':>6}")
    for topic in topics:
        topic_type = get_topic_type(topic)
        average_bandwidth = await get_average_bandwidth(topic, samples=3)
        average_rate = 0
        if average_bandwidth > 0:
            average_rate = await get_average_rate(topic, samples=1)

        print(f"{topic:<50}{topic_type:<50}{average_bandwidth:>8.3f}{average_rate:>6.1f}")

        topic_data.append((topic, topic_type, average_bandwidth, average_rate))

    new_sheet_id = update_spreadsheet(topic_data)

    url = f'https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={new_sheet_id}'
    webbrowser.open(url)


if __name__ == "__main__":
    asyncio.run(main())