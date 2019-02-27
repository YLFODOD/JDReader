import json
import sys
import threading
import time
from json.decoder import JSONDecodeError

import requests
from openpyxl import Workbook


def load_comments_by_page(product_id: int, score: int, sort_type: int, page: int) -> dict:
    """
    :param product_id: the id can be obtained by its url, like https://item.jd.com/2384789.html
    :param score: 1 to 5, 1 is the worst comments
    :param sort_type: 5 sort by date(the latest), 6 default sort type
    :param page: page number
    :return: comments in json
    """
    url = 'https://sclub.jd.com/comment/productPageComments.action'
    # img_prefix = 'http://img30.360buyimg.com/shaidan/' + values from referenceImage
    payload = {
        'callback': 'fetchJSON_comment98vv21549',
        'productId': product_id,
        'score': score,
        'sortType': sort_type,
        'page': page,
        'pageSize': 10,  # change this number makes no difference, more tests required
        'isShadowSku': 0,
        'fold': 1
    }
    r = requests.get(url, params=payload)
    reviews = {}
    if r.ok:
        text = r.text[len(payload['callback']) + 1:-2]  # remove fetchJSON_comment98vv21549(); from the raw content
        try:
            reviews = json.loads(text)
            print('Loading comments from page{}'.format(page))
        except JSONDecodeError as err:
            print(err)
    else:
        print(r.content)

    return reviews


def load_comments(start: int, amount: int, comments: list, cfg) -> None:
    for i in range(start, start + amount):
        r = load_comments_by_page(page=i, **cfg)
        comments.extend(r['comments'])


def run(product_id: int = 2384789, score: int = 1, sort_type: int = 6, workers=10) -> list:
    # Get maxPage from page 0
    page0 = load_comments_by_page(product_id=product_id, score=score, sort_type=sort_type, page=0)
    if len(page0) == 0:
        print('loading failed, please contact the developer for help')

    max_page = page0['maxPage']
    step = max_page // workers
    task = [(i * step, step) if i != (workers - 1) else (i * step, max_page - i * step + 1) for i in range(workers)]
    pool = []
    data = []
    cfg = {
        'product_id': product_id,
        'score': score,
        'sort_type': sort_type,
    }
    for i in task:
        t = threading.Thread(target=load_comments, args=(i[0], i[1], data, cfg), daemon=True)
        pool.append(t)
        t.start()

    for t in pool:
        t.join()

    return data


def save_to_excel(comments: list) -> None:
    wb = Workbook(write_only=True)
    ws = wb.create_sheet(title='comments')
    ws.append(['Date Created', 'Content'])
    for c in comments:
        ws.append([c['creationTime'], c['content']])
    wb.save('comments.xlsx')


if __name__ == '__main__':
    print(time.strftime('%H:%M:%S'))
    _comments = []
    if len(sys.argv) > 1:
        _comments = run(product_id=sys.argv[1])
    else:
        _comments = run()
    save_to_excel(_comments)
    print('The comments was saved in file "comments.xlsx"')
    print(time.strftime('%H:%M:%S'))
