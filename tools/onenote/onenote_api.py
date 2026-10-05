"""OneNote Graph 接口封装：令牌自动续期；建页只发一次（异常时按标题查有没有建成）；建完读回确认。"""
import json, os, time, uuid, urllib.request, urllib.parse
from onenote_auth import CID, TOK, SCOPE
G = 'https://graph.microsoft.com/v1.0/me/onenote'
def tok(): return json.load(open(TOK))['access_token']
def refresh():
    t = json.load(open(TOK))
    d = urllib.parse.urlencode({'client_id': CID, 'grant_type': 'refresh_token', 'refresh_token': t['refresh_token'], 'scope': SCOPE}).encode()
    r = json.load(urllib.request.urlopen(urllib.request.Request('https://login.microsoftonline.com/consumers/oauth2/v2.0/token', data=d), timeout=60))
    r.setdefault('refresh_token', t['refresh_token']); json.dump(r, open(TOK, 'w')); os.chmod(TOK, 0o600)
def _open(method, url, data=None, ctype='application/json', timeout=180):
    for k in range(6):
        r = urllib.request.Request(url, data=data, method=method, headers={'Authorization': 'Bearer ' + tok(), 'Content-Type': ctype})
        try:
            resp = urllib.request.urlopen(r, timeout=timeout); return resp.status, resp.read()
        except urllib.error.HTTPError as e:
            body = e.read().decode(errors='replace')
            if e.code == 401 and k < 2: refresh(); continue
            if e.code in (429, 500, 502, 503, 504) and method == 'GET': time.sleep(5 * (k + 1)); continue
            raise RuntimeError('%s %s' % (e.code, body[:300]))
        except Exception as e:
            if method != 'GET': raise          # 写操作不盲目重发（会产生重复页）
            time.sleep(4)
    raise RuntimeError('重试耗尽 %s' % url)
def req(method, path, body=None, ctype='application/json'):
    url = path if path.startswith('http') else G + path
    data = body.encode() if isinstance(body, str) else (json.dumps(body).encode() if body is not None else None)
    st, t = _open(method, url, data, ctype)
    return json.loads(t) if t else {}
def content(pid, ids=False):
    try:
        return _open('GET', G + '/pages/%s/content%s' % (pid, '?includeIDs=true' if ids else ''))[1].decode()
    except RuntimeError:
        return None
def list_pages(sid):
    return req('GET', '/sections/%s/pages?$select=id,title&$top=100' % sid)['value']
def delete_page(pid):
    try: req('DELETE', '/pages/%s' % pid)
    except RuntimeError as e:
        if '404' not in str(e): raise
def create_page(sid, title, html):
    """建一页并读回确认；读不到就删掉重建（最多 3 次）。返回 page id 或 None。"""
    for att in range(3):
        before = {p['id'] for p in list_pages(sid) if p['title'] == title}
        try:
            pid = json.loads(_open('POST', G + '/sections/%s/pages' % sid, html.encode(), 'text/html')[1])['id']
        except Exception:
            time.sleep(20)
            new = [p['id'] for p in list_pages(sid) if p['title'] == title and p['id'] not in before]
            if not new: continue
            pid = new[0]
        for w in range(8):
            time.sleep(8); h = content(pid)
            if h:
                if '<title></title>' in h or '<title>' not in h:
                    req('PATCH', '/pages/%s/content' % pid, [{'target': 'title', 'action': 'replace', 'content': title}])
                return pid
        delete_page(pid)
    return None
def patch_with_image(pid, commands, img_name, img_bytes, mime='image/png'):
    b = '----b' + uuid.uuid4().hex
    body = (('--%s\r\nContent-Disposition: form-data; name="Commands"\r\nContent-Type: application/json\r\n\r\n%s\r\n' % (b, json.dumps(commands))).encode()
            + ('--%s\r\nContent-Disposition: form-data; name="%s"\r\nContent-Type: %s\r\n\r\n' % (b, img_name, mime)).encode() + img_bytes + ('\r\n--%s--\r\n' % b).encode())
    return _open('PATCH', G + '/pages/%s/content' % pid, body, 'multipart/form-data; boundary=' + b, timeout=300)[0]
