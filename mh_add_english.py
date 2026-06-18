"""
mh_add_english.py — add an English display label to every node of
config/myheritage_categories.json.

Per the user: KEEP the Russian tree; for each Russian name (1) search VERY
thoroughly for its exact English equivalent in results/MH-English.docx, and ONLY
if not found (2) translate it ourselves (slightly inaccurate is fine — the Russian
is shown next to it in the GUI).

Matching uses three language-independent / dictionary signals:
  • the set of 4-digit years in the name (1538-1900 etc.) — strong, language-neutral;
  • transliterated place names (Бристоль→bristol, Йоркшир→yorkshire);
  • a geo + genealogy term dictionary (Дания→Denmark, рождения→births, церков→church…).
A name is matched to the EN-docx entry with the same year-set and the highest token
overlap. If nothing scores high enough, ru_to_en() builds a term-by-term translation
(the same dictionaries) as the fallback.

Output: config/myheritage_categories.json rewritten as
  { ru_label: { "en": <english>, "children": { … } } }
(ru stays the key/structure; en is display + English-site search).
"""
from __future__ import annotations
import difflib
import json
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

from docx import Document

_HERE = Path(__file__).resolve().parent
_JSON = _HERE / "config" / "myheritage_categories.json"
_ENDOCX = _HERE / "results" / "MH-English.docx"
_NUMONLY = re.compile(r'^[\d\s.,+]+$')

# ── transliteration (for English place names written in Cyrillic) ──────────────
_TR = {'а':'a','б':'b','в':'v','г':'g','д':'d','е':'e','ё':'e','ж':'zh','з':'z',
       'и':'i','й':'i','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r',
       'с':'s','т':'t','у':'u','ф':'f','х':'h','ц':'c','ч':'ch','ш':'sh','щ':'sch',
       'ъ':'','ы':'y','ь':'','э':'e','ю':'yu','я':'ya'}
def _translit(s): return ''.join(_TR.get(c, c) for c in s.lower())

# ── geo names (translated, not transliterated) ─────────────────────────────────
GEO = {
 'дания':'Denmark','англия':'England','германия':'Germany','франция':'France',
 'италия':'Italy','испания':'Spain','швеция':'Sweden','норвегия':'Norway',
 'нидерланды':'Netherlands','голландия':'Holland','бельгия':'Belgium',
 'австрия':'Austria','австро':'Austria','венгрия':'Hungary','чехия':'Czechia',
 'словакия':'Slovakia','польша':'Poland','россия':'Russia','украина':'Ukraine',
 'ирландия':'Ireland','шотландия':'Scotland','уэльс':'Wales','греция':'Greece',
 'финляндия':'Finland','швейцария':'Switzerland','канада':'Canada',
 'австралия':'Australia','зеландия':'Zealand','аргентина':'Argentina',
 'бразилия':'Brazil','мексика':'Mexico','чили':'Chile','уругвай':'Uruguay',
 'никарагуа':'Nicaragua','гватемала':'Guatemala','сальвадор':'Salvador',
 'боливия':'Bolivia','венесуэла':'Venezuela','ямайка':'Jamaica',
 'барбадос':'Barbados','багамы':'Bahamas','индия':'India','израиль':'Israel',
 'латвия':'Latvia','литва':'Lithuania','эстония':'Estonia','румыния':'Romania',
 'португалия':'Portugal','исландия':'Iceland','колумбия':'Colombia',
 'сша':'United States','соединенные':'United','штаты':'States','королевство':'Kingdom',
 'британия':'Britain','великобритания':'Great Britain','соединенное':'United',
 'южная':'South','северная':'North','западная':'West','восточная':'East','остров':'Island',
 'африка':'Africa','пруссия':'Prussia','бавария':'Bavaria','саксония':'Saxony',
 'мекленбург':'Mecklenburg','шверин':'Schwerin','бранденбург':'Brandenburg',
 'познань':'Posen','померания':'Pomerania','гессен':'Hesse','баден':'Baden',
 'вестфалия':'Westphalia','рейн':'Rhine','каталония':'Catalonia',
 'калифорния':'California','техас':'Texas','огайо':'Ohio','индиана':'Indiana',
 'иллинойс':'Illinois','миннесота':'Minnesota','мичиган':'Michigan',
 'висконсин':'Wisconsin','миссури':'Missouri','канзас':'Kansas',
 'небраска':'Nebraska','айова':'Iowa','аризона':'Arizona','невада':'Nevada',
 'юта':'Utah','монтана':'Montana','вермонт':'Vermont','мэн':'Maine',
 'мэриленд':'Maryland','вирджиния':'Virginia','каролина':'Carolina',
 'джорджия':'Georgia','флорида':'Florida','алабама':'Alabama',
 'миссисипи':'Mississippi','луизиана':'Louisiana','арканзас':'Arkansas',
 'теннесси':'Tennessee','кентукки':'Kentucky','оклахома':'Oklahoma',
 'колорадо':'Colorado','айдахо':'Idaho','вашингтон':'Washington',
 'орегон':'Oregon','гавайи':'Hawaii','коннектикут':'Connecticut',
 'массачусетс':'Massachusetts','пенсильвания':'Pennsylvania','джерси':'Jersey',
 'род':'Rhode','айленд':'Island','гэмпшир':'Hampshire','дакота':'Dakota',
 'йоркшир':'Yorkshire','ланкашир':'Lancashire','чешир':'Cheshire',
 'дербишир':'Derbyshire','эссекс':'Essex','дорсет':'Dorset','норфолк':'Norfolk',
 'кент':'Kent','корнуолл':'Cornwall','бристоль':'Bristol','дублин':'Dublin',
 'вена':'Vienna','квинсленд':'Queensland','тасмания':'Tasmania','квебек':'Quebec',
 'онтарио':'Ontario','манитоба':'Manitoba','саскачеван':'Saskatchewan',
}

# the 14 top categories — exact English (these ARE the EN-docx L1 names).
L1_EN = {
 "Школы и университеты": "Schools & Universities",
 "Перепись и списки избирателей": "Census & Voter Lists",
 "Справочники, путеводители и ссылки": "Directories, Guides & References",
 "Публичные очеты": "Public Records",
 "Реестры рождения, браков и смерти": "Birth, Marriage & Death",
 "Истории, мемуары и биографии": "Histories, Memories & Biographies",
 "Карты": "Maps", "Фото": "Photos", "Вооруженные силы": "Military",
 "Правительство, земля, суды и завещания": "Government, Land, Court & Wills",
 "Книги и публикации": "Books & Publications", "Газеты": "Newspapers",
 "Иммиграция и путешествия": "Immigration & Travel",
 "Семейные деревья": "Family Trees",
}

# ── genealogy / record terms (stem → English) ─────────────────────────────────
TERMS = [
 ('бракосочетани','marriage'),('расторжени','divorce'),('утверждени','probate'),('рождени','birth'),
 ('рожден','birth'),('крещени','baptism'),('брак','marriage'),('развод','divorce'),
 ('смерт','death'),('погребени','burial'),('захорони','burial'),('захорон','burial'),
 ('похорон','funeral'),('некролог','obituary'),('кладбищ','cemetery'),
 ('церков','church'),('синагог','synagogue'),('метрическ','parish register'),
 ('переписн','census'),('перепис','census'),('избирател','voter'),
 ('указател','index'),('индекс','index'),('реестр','records'),('записе','records'),
 ('земельн','land'),('заключённ','prisoner'),('заключенн','prisoner'),('юридическ','legal'),
 ('запис','records'),('документ','record'),('призывн','draft'),('призыв','draft'),
 ('наем','enlistment'),('служб','service'),('военн','military'),('армии','army'),
 ('флот','navy'),('иммиграци','immigration'),('эмиграци','emigration'),
 ('пассажир','passenger'),('паспорт','passport'),('натурализаци','naturalization'),
 ('гражданств','citizenship'),('завещани','will'),('суд','court'),('земл','land'),
 ('недвижимост','real estate'),('правительств','government'),('газет','newspaper'),
 ('справочник','directory'),('телефонн','phone'),('путеводител','guide'),
 ('ссылк','reference'),('студент','student'),('университет','university'),
 ('колледж','college'),('школьн','school'),('школ','school'),('ученик','pupil'),
 ('ежегодник','yearbook'),('выпускник','alumni'),('биографи','biography'),
 ('истори','history'),('мемуар','memoir'),('карт','map'),('фотографи','photo'),
 ('фото','photo'),('семейн','family'),('дерев','tree'),('дубликат','duplicate'),
 ('книг','book'),('публикаци','publication'),('акт','certificate'),
 ('гражданск','civil'),('состояни','status'),('регистраци','registration'),
 ('списк','list'),('реквизит','reference'),('консистори','consistory'),
 ('приход','parish'),('католическ','catholic'),('реформатск','reformed'),
 ('протестант','protestant'),('православн','orthodox'),('еврейск','jewish'),
 ('население','population'),('населени','population'),('владельц','owners'),
 ('пенси','pension'),('ветеран','veteran'),('пособи','benefits'),('заявлени','applications'),
 ('требовани','claims'),('фонд','fund'),('опис','inventory'),('дело','file'),
 ('известн','notable'),('знаменит','famous'),('замечательн','notable'),
 ('новорожденн','newborn'),('умерш','deceased'),('бизнес','business'),
 ('предприяти','business'),('коллекци','collection'),('сборник','collection'),
 ('реестры жизни','life records'),('жизни','life'),('другие','other'),
 ('другое','other'),('прочее','other'),('всего мира','worldwide'),
 ('всему миру','worldwide'),('часть','part'),('том','volume'),
 ('национальн','national'),('федеральн','federal'),('окружн','county'),('округ','county'),
 ('город','city'),('городск','city'),('штат','state'),('провинци','province'),
 ('район','region'),('област','region'),('губерни','province'),
]

_STOP_EN = set("the of and a in to for by on at with de la el los las и в по о с от на за к".split())
def _stem(w):
    for suf in ("ies", "es", "s"):
        if len(w) > 4 and w.endswith(suf):
            return w[:-len(suf)] + ("y" if suf == "ies" else "")
    return w
def _deacc(s):
    return ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
def _years(s): return frozenset(re.findall(r'\b(1[5-9]\d\d|20\d\d)\b', s))
def _ten(s):
    """Comparable token set: deaccented, lowercased, stemmed. Keeps small numbers like a
    «Part 5» (so part 5 ≠ part 4) but NOT 4-digit years (those go in the year-set)."""
    out = set()
    for w in re.findall(r'[a-z0-9]+', _deacc(s).lower()):
        if w in _STOP_EN:
            continue
        if w.isdigit():
            if not re.fullmatch(r'1[5-9]\d\d|20\d\d', w):
                out.add(w)
        elif len(w) > 2:
            out.add(_stem(w))
    return out
def _fuzov(rt, et):
    """How many RU tokens have an equal or near-equal (≥0.85) EN token (place-name drift
    like Оребро→orebro≈örebro→orebro)."""
    n = 0
    for r in rt:
        if r in et:
            n += 1
        elif any(len(e) > 3 and difflib.SequenceMatcher(None, r, e).ratio() >= 0.85 for e in et):
            n += 1
    return n

def _geo(word):
    """Geo translation tolerant of Russian case endings (Канады→Canada): match the
    geo key against the word's stem (key minus its last 1-2 letters)."""
    w = word.lower()
    if w in GEO:
        return GEO[w]
    for k, v in GEO.items():
        if len(k) >= 5 and w.startswith(k[:-1]):
            return v
    return None

def ru_to_en(ru: str) -> str:
    """Term-by-term English translation (fallback when no EN-docx match)."""
    s = ' ' + ru.lower() + ' '
    for stem, en in TERMS:                    # longest-ish stems first roughly
        s = re.sub(stem + r'[а-яё]*', ' ' + en + ' ', s)
    out_words = []
    for w in re.findall(r"[а-яёa-z0-9'’/.-]+", s):
        lw = w.lower()
        if lw in ("и", "а"):
            out_words.append("&")
        elif _geo(lw):
            out_words.append(_geo(lw))
        elif re.fullmatch(r"[a-z0-9'’/.-]+", lw) or re.fullmatch(r'\d{4}([–-]\d{4})?', w):
            out_words.append(w)            # already latin / year
        elif re.search(r'[а-яё]', lw):
            out_words.append(_translit(w))  # place name → translit
        else:
            out_words.append(w)
    res = re.sub(r'\s+', ' ', ' '.join(out_words)).strip(' ,;&')
    return res[:1].upper() + res[1:] if res else ru

def _tokens_ru_as_en(ru: str) -> set:
    return _ten(ru_to_en(ru))

def main():
    en_doc = Document(str(_ENDOCX))
    en_names = []
    for p in en_doc.paragraphs:
        t = p.text.strip()
        if not t or t.lower() == "all collections" or _NUMONLY.match(t):
            continue
        if "<br>" in t or len(t) > 160:
            continue
        en_names.append(t)
    en_names = list(dict.fromkeys(en_names))
    EN = [(e, _ten(e), _years(e)) for e in en_names]
    yidx, tidx = defaultdict(set), defaultdict(set)   # year/token → candidate indices
    for i, (e, et, ey) in enumerate(EN):
        for y in ey:
            yidx[y].add(i)
        for w in et:
            tidx[w].add(i)

    def ranked(ru):
        """Ranked EN candidates for `ru` (best first). Candidates share a year (or, for
        no-year names, a token); score = fuzzy Jaccard, years weighted ×2."""
        ry = _years(ru); rt = _ten(ru_to_en(ru))
        cand = set()
        for y in ry:
            cand |= yidx[y]
        if not ry:
            for w in rt:
                cand |= tidx.get(w, set())
        if not cand:
            cand = set(range(len(EN)))
        scored = []
        for i in cand:
            e, et, ey = EN[i]
            yo = len(ry & ey); to = _fuzov(rt, et)
            inter = to + 2 * yo
            union = len(rt) + 2 * len(ry) + len(et) + 2 * len(ey) - inter
            scored.append((inter / max(union, 1), i))
        scored.sort(reverse=True)
        return scored[:12]

    tree = json.loads(_JSON.read_text("utf-8"))
    ru_names = []
    def collect(t):
        for ru, v in t.items():
            ru_names.append(ru)
            if v.get("children"):
                collect(v["children"])
    collect(tree)

    def is_latin(s):           # already-English names (book titles, «BillionGraves»…)
        return not re.search(r'[а-яё]', s, re.I)

    # manual English overrides (user-supplied exact titles for names the EN docx lacks)
    ovr_path = _HERE / "config" / "mh_en_overrides.json"
    OVR = json.loads(ovr_path.read_text("utf-8")) if ovr_path.exists() else {}

    assign, score_of = {}, {}
    for ru in set(ru_names):
        if ru in OVR:                             # user-supplied exact English — highest priority
            assign[ru], score_of[ru] = OVR[ru], 1.0
        elif is_latin(ru):                        # already English → keep as-is (exact)
            assign[ru], score_of[ru] = ru, 1.0
        elif ru in L1_EN:                         # the 14 top categories — exact
            assign[ru], score_of[ru] = L1_EN[ru], 1.0
        else:
            c = ranked(ru)
            if c and c[0][0] >= 0.2:              # confident EN-docx match
                assign[ru], score_of[ru] = EN[c[0][1]][0], c[0][0]
            else:                                 # not findable in EN docx → translate
                assign[ru], score_of[ru] = ru_to_en(ru), (c[0][0] if c else 0.0)

    low = sorted(((score_of[ru], ru, assign[ru]) for ru in assign), key=lambda x: x[0])
    n_low = sum(1 for s, _, _ in low if s < 0.34)

    def walk(t):
        new = {}
        for ru, v in t.items():
            node = {"en": assign.get(ru, ru_to_en(ru))}
            if v.get("children"):
                node["children"] = walk(v["children"])
            new[ru] = node
        return new

    out = walk(tree)
    _JSON.write_text(json.dumps(out, ensure_ascii=False, indent=2), "utf-8")
    print(f"nodes: {len(ru_names)} unique: {len(assign)} | low-confidence (<0.34): {n_low}")
    print("--- lowest-confidence (verify these): ---")
    for sc, ru, en in low[:30]:
        print(f"  [{sc:.2f}] {ru[:38]}  →  {en[:38]}")


if __name__ == "__main__":
    main()


if __name__ == "__main__":
    main()
