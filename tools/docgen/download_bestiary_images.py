"""Загрузка иллюстраций для бестиария (docs/04_Бестиарий_Берсерк_референсы.docx).

Изображения — собственность правообладателей «Берсерка», берутся с Berserk Wiki
только как внутренний референс. В игре и в публичных материалах не использовать.

Запуск из корня проекта:
    python tools/docgen/download_bestiary_images.py
    python tools/docgen/build.py
"""
import sys
import time
import urllib.request
from pathlib import Path

OUT = Path(__file__).resolve().parents[2] / "docs" / "bestiary_img"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130 Safari/537.36",
    "Referer": "https://berserk.fandom.com/",
}
IMAGES = [
    ("Void", "https://static.wikia.nocookie.net/berserk/images/7/74/Manga_E76_Void.png"),
    ("Slan", "https://static.wikia.nocookie.net/berserk/images/1/1a/Manga_E75_Slan.png"),
    ("Ubik", "https://static.wikia.nocookie.net/berserk/images/b/b6/Ubik_Manga.jpg"),
    ("Conrad", "https://static.wikia.nocookie.net/berserk/images/f/f1/Conrad_Manga.jpg"),
    ("Griffith", "https://static.wikia.nocookie.net/berserk/images/8/8a/Manga_E86_Femto.png"),
    ("Zodd", "https://static.wikia.nocookie.net/berserk/images/c/c0/BTCG_Zodd_Full_Body.png"),
    ("Grunbeld", "https://static.wikia.nocookie.net/berserk/images/4/40/Grunbeld_Profile_Miura.png"),
    ("Locus", "https://static.wikia.nocookie.net/berserk/images/8/88/Locus_Manga.jpg"),
    ("Irvine", "https://static.wikia.nocookie.net/berserk/images/d/d4/Irvine_Manga.jpg"),
    ("Rakshas", "https://static.wikia.nocookie.net/berserk/images/d/dc/Rakshas_Manga.jpg"),
    ("Rosine", "https://static.wikia.nocookie.net/berserk/images/f/f7/Rosine_Manga.jpg"),
    ("The_Count", "https://static.wikia.nocookie.net/berserk/images/e/e1/Manga_E0-3_The_Count.png"),
    ("Snake_Lord", "https://static.wikia.nocookie.net/berserk/images/6/6a/Manga_E0-1_Snake_Lord_Armored.png"),
    ("Wyald", "https://static.wikia.nocookie.net/berserk/images/9/95/Wyald_Manga.jpg"),
    ("Ganishka", "https://static.wikia.nocookie.net/berserk/images/b/b3/Manga_V32_Ganishka.png"),
    ("Egg_of_the_Perfect_World", "https://static.wikia.nocookie.net/berserk/images/0/02/Egg-Shaped_Apostle.jpg"),
    ("Borkoff", "https://static.wikia.nocookie.net/berserk/images/2/2d/Borkoff_apostle.png"),
    ("Balzac", "https://static.wikia.nocookie.net/berserk/images/8/84/Balzac_Apostle.jpg"),
    ("Vlad_Tepes", "https://static.wikia.nocookie.net/berserk/images/b/b9/Vlad_Tepes_Manga.jpg"),
    ("Unidentified_Roaming_Apostle", "https://static.wikia.nocookie.net/berserk/images/4/47/Keeper_Demon.png"),
    ("Pseudo-Apostle", "https://static.wikia.nocookie.net/berserk/images/1/1e/Berserk_v2_p139.png"),
    ("Beast_of_Darkness", "https://static.wikia.nocookie.net/berserk/images/f/fc/Manga_E226_Beast_of_Darkness.png"),
    ("Dead_Spirit", "https://static.wikia.nocookie.net/berserk/images/2/2a/Serpico_surrounded_by_demons.png"),
    ("Troll", "https://static.wikia.nocookie.net/berserk/images/1/19/Trolls.jpg"),
    ("Ogre", "https://static.wikia.nocookie.net/berserk/images/7/75/Ogre.jpg"),
    ("Kelpie", "https://static.wikia.nocookie.net/berserk/images/3/37/Kelpie.jpg"),
    ("Merrow", "https://static.wikia.nocookie.net/berserk/images/8/80/Merrow.png"),
    ("Sea_God", "https://static.wikia.nocookie.net/berserk/images/5/58/Sea_God_Mouth.png"),
    ("Daka", "https://static.wikia.nocookie.net/berserk/images/b/bc/Daka.jpg"),
    ("Pishacha", "https://static.wikia.nocookie.net/berserk/images/c/c1/Pishacha_Tiger.jpg"),
    ("Makara", "https://static.wikia.nocookie.net/berserk/images/4/44/Manga_E240_Makara.png"),
    ("Garuda", "https://static.wikia.nocookie.net/berserk/images/7/78/Garuda.png"),
    ("Kundalini", "https://static.wikia.nocookie.net/berserk/images/0/07/Kundalini.jpg"),
    ("Chimimoryo", "https://static.wikia.nocookie.net/berserk/images/6/65/Manga_E215_Chimimoryo.png"),
    ("Skull_Knight", "https://static.wikia.nocookie.net/berserk/images/6/6f/BTCG_Skull_Knight.png"),
    ("Golem", "https://static.wikia.nocookie.net/berserk/images/6/64/Golems.jpg"),
    ("Apostle", "https://static.wikia.nocookie.net/berserk/images/3/3e/Manga_E73_Apostles.png"),
    ("Beherit", "https://static.wikia.nocookie.net/berserk/images/b/b2/Manga_E0-4_Beherit.png"),
    ("Elemental", "https://static.wikia.nocookie.net/berserk/images/3/3f/Sylphs.png"),
    ("Evil", "https://static.wikia.nocookie.net/berserk/images/2/24/Manga_E83_Griffith_fashioning_a_proper_form.jpg"),
]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    OUT.mkdir(parents=True, exist_ok=True)
    ok = 0
    for name, url in IMAGES:
        ext = Path(url).suffix.lower().replace(".jpeg", ".jpg")
        target = OUT / (name + ext)
        if target.exists():
            ok += 1
            continue
        try:
            data = urllib.request.urlopen(urllib.request.Request(url, headers=HEADERS), timeout=60).read()
            target.write_bytes(data)
            ok += 1
            print("ok  ", target.name)
        except Exception as e:
            print("ошибка", name, e)
        time.sleep(0.5)
    print(f"готово: {ok} из {len(IMAGES)}")


if __name__ == "__main__":
    main()
