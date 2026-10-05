#!/usr/bin/env python3
"""OneNote 设备码登录：python3 onenote_auth.py
屏幕显示网址和代码，用户自己在浏览器登录、同意「读写 OneNote 笔记本」。
令牌存 ~/.config/flight-onenote/tok.json（权限 600，不进仓库）；之后由 onenote_api 自动续期。"""
import json, os, time, urllib.request, urllib.parse
CID = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft Graph Command Line Tools（微软官方公共客户端）
AUTH = 'https://login.microsoftonline.com/consumers/oauth2/v2.0'
SCOPE = 'Notes.ReadWrite offline_access'
TOK = os.path.expanduser('~/.config/flight-onenote/tok.json')
def post(url, data):
    for k in range(5):
        try:
            return json.load(urllib.request.urlopen(urllib.request.Request(url, data=urllib.parse.urlencode(data).encode()), timeout=30))
        except urllib.error.HTTPError as e:
            return json.load(e)
        except Exception as e:
            print('网络重试', k + 1, str(e)[:60], flush=True); time.sleep(3)
    return {'error': 'authorization_pending'}
if __name__ == '__main__':
    os.makedirs(os.path.dirname(TOK), exist_ok=True)
    d = post(AUTH + '/devicecode', {'client_id': CID, 'scope': SCOPE})
    print(d.get('message') or d, flush=True)
    t0 = time.time()
    while time.time() - t0 < d.get('expires_in', 900):
        time.sleep(d.get('interval', 5))
        r = post(AUTH + '/token', {'grant_type': 'urn:ietf:params:oauth:grant-type:device_code', 'client_id': CID, 'device_code': d['device_code']})
        if 'access_token' in r:
            json.dump(r, open(TOK, 'w')); os.chmod(TOK, 0o600); print('OK 已获得令牌'); break
        if r.get('error') != 'authorization_pending':
            print('ERR', r.get('error'), r.get('error_description', '')[:200]); break
