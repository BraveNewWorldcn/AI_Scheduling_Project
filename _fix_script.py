
import sys

path = r"f:\AI\AI_Scheduling_Project\customer_agent.py"

with open(path, 'rb') as f:
    data = f.read()

# Fix 2a: chat_ids init + max_processed_ids + last_token_refresh
old = b'    processed_ids: set = set()\r\n    poll_interval = 2.0\r\n    last_chat_refresh = datetime.min.replace(tzinfo=CST)'
new = (
    b'    chat_ids: List[str] = []\r\n'
    b'    processed_ids: set = set()\r\n'
    b'    max_processed_ids = 10000\r\n'
    b'    poll_interval = 2.0\r\n'
    b'    last_chat_refresh = datetime.min.replace(tzinfo=CST)\r\n'
    b'    last_token_refresh = datetime.now(CST)'
)
if old in data:
    data = data.replace(old, new)
    print('Fix 2a+2b OK')
else:
    print('ERROR 2a: pattern not found')

# Fix 2c: Add token refresh + use bot_open_id in _list_chat_messages call
old2 = b'            for chat_id in chat_ids:\r\n                messages = _list_chat_messages(token, chat_id)'
new2 = (
    b'            if (now - last_token_refresh).total_seconds() > 5400:\r\n'
    b'                token = _get_bot_token()\r\n'
    b'                last_token_refresh = now\r\n'
    b'                print(f"[{now.strftime(\"%H:%M:%S\")}] Token \xe5\xb7\xb2\xe5\x88\xb7\xe6\x96\xb0")\r\n'
    b'\r\n'
    b'            for chat_id in chat_ids:\r\n'
    b'                messages = _list_chat_messages(token, chat_id, bot_open_id)'
)
if old2 in data:
    data = data.replace(old2, new2)
    print('Fix 2c OK')
else:
    print('ERROR 2c: pattern not found')
    idx = data.find(b'_list_chat_messages(token, chat_id)')
    print(f'Found at {idx}')
    if idx >= 0:
        print(repr(data[idx-40:idx+60]))

# Fix 3: processed_ids cleanup
old3 = b'            for chat_id in chat_ids:\r\n                messages = _list_chat_messages(token, chat_id, bot_open_id)'
new3 = (
    b'            if len(processed_ids) > max_processed_ids:\r\n'
    b'                keep = max_processed_ids // 2\r\n'
    b'                processed_ids = set(list(processed_ids)[-keep:])\r\n'
    b'                print(f"[{now.strftime(\"%H:%M:%S\")}] processed_ids \xe5\xb7\xb2\xe6\xb8\x85\xe7\x90\x86 (\xe4\xbf\x9d\xe7\x95\x99 {keep} \xe6\x9d\xa1)")\r\n'
    b'\r\n'
    b'            for chat_id in chat_ids:\r\n'
    b'                messages = _list_chat_messages(token, chat_id, bot_open_id)'
)
if old3 in data:
    data = data.replace(old3, new3)
    print('Fix 3 OK')
else:
    print('ERROR 3: pattern not found for cleanup')
    idx = data.find(b'_list_chat_messages(token, chat_id, bot_open_id)')
    print(f'Found bot_open_id call at {idx}')

# Fix 4: traceback in exception
old4 = b'            print(f"[X] \xe8\xbd\xae\xe8\xaf\xa2\xe5\xbc\x82\xe5\xb8\xb8 ({datetime.now(CST).strftime(\"%H:%M:%S\")}): {e}")\r\n            time.sleep(5)'
new4 = b'            print(f"[X] \xe8\xbd\xae\xe8\xaf\xa2\xe5\xbc\x82\xe5\xb8\xb8 ({datetime.now(CST).strftime(\"%H:%M:%S\")}): {e}")\r\n            import traceback\r\n            traceback.print_exc()\r\n            time.sleep(5)'
if old4 in data:
    data = data.replace(old4, new4)
    print('Fix 4 OK')
else:
    print('ERROR 4: pattern not found')

with open(path, 'wb') as f:
    f.write(data)
print('Done saving')
