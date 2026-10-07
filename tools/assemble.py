# -*- coding: utf-8 -*-
"""standalone_template.html + seed.json -> 営業資料ファインダー.html

ひな形（標準HTML。データ部分が目印のまま）に初期データを差し込んで、
配布用の単体HTMLを作る。生成したHTMLはひな形を内側に抱えるので、
ブラウザの「HTMLに保存」からも同じ形で書き出せる。
"""
import json

TPL = "standalone_template.html"
SEED = "seed.json"
OUT = "営業資料ファインダー.html"


def esc(value):
    """JS/JSONに埋め込める表記にする。閉じタグの並びは <\\/ に逃がして、
    スクリプトが途中で終わったと誤認されないようにする（JSONでは \\/ は / に戻る）。"""
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")


template = open(TPL, encoding="utf-8").read()
seed = json.load(open(SEED, encoding="utf-8"))

assert template.count('"__TEMPLATE__"') == 1, "ひな形の目印（TEMPLATE）が1つではない"
assert template.count("__SEED_OBJ__") == 1, "ひな形の目印（SEED）が1つではない"

final = template.replace('"__TEMPLATE__"', esc(template), 1)
final = final.replace("__SEED_OBJ__", esc(seed), 1)

n = final.count("</script>")
assert n == 1, "スクリプトの閉じタグが %d 個ある（1個でなければページが途中で切れる）" % n
assert json.loads(esc(template)) == template, "埋め込んだひな形が元に戻らない"
assert json.loads(esc(seed)) == seed, "埋め込んだ初期データが元に戻らない"
assert "var SEED = {" in final, "初期データがオブジェクトとして埋め込まれていない"

open(OUT, "w", encoding="utf-8").write(final)
print("wrote %s / %d chars / 資料 %d件 / 定型文 %d件 / セット %d件"
      % (OUT, len(final), len(seed.get("items", [])),
         len(seed.get("blurbs", [])), len(seed.get("sets", []))))
