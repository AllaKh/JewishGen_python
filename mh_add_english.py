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
 'род':'Rhode','айленд':'Island','род-айленд':'Rhode Island','юта':'Utah',
 'нью-йорк':'New York','нью-джерси':'New Jersey','нью-гэмпшир':'New Hampshire',
 'гэмпшир':'Hampshire','дакота':'Dakota',
 'кальвадос':'Calvados','рона':'Rhône','сена':'Seine','сены':'Seine','роны':'Rhône',
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
 ('смерти','deaths'),('смертност','mortality'),('смерт','death'),
 ('погребени','burial'),('захоронени','burials'),('захорони','burial'),('захорон','burial'),
 ('окружн','county'),('округ','county'),('призывник','conscripts'),('иммигрирующ','immigrating'),
 ('немц','Germans'),('надгробн','headstone'),('ополчени','militia'),('увольнени','discharge'),
 ('аттестат','certificate'),('приход','parish'),('извещени','notices'),
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
def _year_ok(ry, ey):
    """True if the RU year-set is compatible with the EN candidate's year-set: either no
    years on a side, an exact shared year, or an RU year that falls inside the EN's
    [min,max] span (so «1850» still matches a «1837-2005» range title). This blocks the
    census/electoral bug where «Перепись Швеции 1880» grabbed «1940 Sweden Census»."""
    if not ry or not ey:
        return True
    if ry & ey:
        return True
    lo, hi = min(ey), max(ey)
    return any(lo <= y <= hi for y in ry)
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

# Whole multi-word phrases → English, applied BEFORE the word-by-word pass so common
# genealogy collocations don't come out as word salad («Карточки воинского учёта» →
# «map voinskogo ucheta»). Longest first. The English here is already Latin, so the
# word loop keeps it verbatim.
PHRASES = [
    ("карточки воинского учёта", "Military Registration Cards"),
    ("карточки воинского учета", "Military Registration Cards"),
    ("воинского учёта", "Military Registration"),
    ("воинского учета", "Military Registration"),
    ("реестр записей регистрации смерти", "Death Registration Records"),
    ("записей регистрации смерти", "Death Registration"),
    ("записи регистрации смерти", "Death Registration Records"),
    ("приходские записи о смертях и захоронениях", "Parish Death and Burial Records"),
    ("приходские записи о рождениях и крещениях", "Parish Birth and Baptism Records"),
    ("приходские записи", "Parish Records"),
    ("записи о смерти и погребении", "Death and Burial Records"),
    ("акты гражданского состояния", "Civil Status Records"),
    ("гражданского состояния", "Civil Status"),
    ("записи о рождениях", "Birth Records"),
    ("записи о рождении", "Birth Records"),
    ("записи о смерти", "Death Records"),
    ("записи о браке", "Marriage Records"),
    ("записи о браках", "Marriage Records"),
    ("свидетельства о смерти", "Death Certificates"),
    ("свидетельства о рождении", "Birth Certificates"),
    ("акты о смерти", "Death Certificates"),
    ("списки избирателей", "Electoral Rolls"),
    ("списки пассажиров", "Passenger Lists"),
    ("перепись населения", "Census"),
    ("переписи населения", "Census"),
]

def ru_to_en(ru: str) -> str:
    """Term-by-term English translation (fallback when no EN-docx match)."""
    s = ' ' + ru.lower() + ' '
    for ru_ph, en_ph in PHRASES:              # whole-phrase pass first
        s = re.sub(r'(?<![а-яё])' + re.escape(ru_ph) + r'(?![а-яё])',
                   ' ' + en_ph + ' ', s)
    for stem, en in TERMS:                    # longest-ish stems first roughly
        s = re.sub(stem + r'[а-яё]*', ' ' + en + ' ', s)
    PREP = {"и": "&", "а": "&", "в": "in", "во": "in", "от": "from", "для": "for",
            "по": "by", "о": "", "об": "", "на": "", "с": "", "со": "", "у": ""}
    out_words = []
    for w in re.findall(r"[A-Za-zА-ЯЁа-яё0-9'’/.-]+|,", s):  # keep case + commas
        if w == ",":
            if out_words and out_words[-1] != ",":
                out_words.append(",")
            continue
        lw = w.lower()
        if lw in PREP:
            if PREP[lw]:
                out_words.append(PREP[lw])
        elif lw in ("год", "года", "годов", "году", "гг", "г", "штат", "штата", "штате"):
            continue                        # year-noise / «штат» filler → drop
        elif _geo(lw):
            out_words.append(_geo(lw))
        elif re.fullmatch(r"[A-Za-z0-9'’/.-]+", w) or re.fullmatch(r'\d{4}([–-]\d{4})?', w):
            out_words.append(w)            # already latin / year
        elif re.search(r'[а-яё]', lw):
            t = _translit(w)               # place name → translit (capitalised)
            out_words.append(t[:1].upper() + t[1:] if t else t)
        else:
            out_words.append(w)
    res = re.sub(r'\s+', ' ', ' '.join(out_words)).replace(' ,', ',').strip(' ,;&')
    # «county Allen» → «Allen County», «parish Leicestershire» → «Leicestershire Parish»
    res = re.sub(r'\b(county|parish)\s+([A-Z][\w’-]+)', r'\2 ' + r'\1', res)
    res = re.sub(r'\b(county|parish)\b', lambda m: m.group(1).capitalize(), res)
    # «Census Sweden 1880» → «1880 Sweden Census» (so the right year leads, like the site)
    res = re.sub(r'^Census\s+([A-Z][\w’ -]*?)\s+(\d{4})$', r'\2 \1 Census', res)
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

    assign, score_of, taken = {}, {}, set()
    fuzzy = []
    for ru in set(ru_names):
        # year guard: a few summary-table blocks (census / electoral rolls) zipped
        # RU↔EN one row off, pairing «…1880» with «1940 … Census». Reject an override
        # whose year clashes with the RU year → it falls through to fuzzy/dict, which
        # rebuilds the correct year. (The file is left intact — nothing deleted.)
        if ru in OVR and _year_ok(_years(ru), _years(OVR[ru])):  # exact English, top priority
            assign[ru], score_of[ru] = OVR[ru], 1.0; taken.add(OVR[ru])
        elif is_latin(ru):                        # already English → keep as-is (exact)
            assign[ru], score_of[ru] = ru, 1.0; taken.add(ru)
        elif ru in L1_EN:                         # the 14 top categories — exact
            assign[ru], score_of[ru] = L1_EN[ru], 1.0; taken.add(L1_EN[ru])
        else:
            fuzzy.append(ru)
    # UNIQUE greedy: confident names grab their EN first; each EN used once, so a wrong
    # match can't reuse an EN that belongs to another node (the «Romania» reuse bug).
    rk = {ru: ranked(ru) for ru in fuzzy}
    for ru in sorted(fuzzy, key=lambda r: -(rk[r][0][0] if rk[r] else 0)):
        ry = _years(ru)
        picked, psc = None, 0.0
        for sc, i in rk[ru]:
            if sc < 0.2:
                break
            e, _et, ey = EN[i]
            if not _year_ok(ry, ey):          # never assign a different-year title
                continue
            if e not in taken:
                picked, psc = e, sc; taken.add(e); break
        if picked is None:                        # nothing confident & free → translate
            picked, psc = ru_to_en(ru), (rk[ru][0][0] if rk[ru] else 0.0)
        assign[ru], score_of[ru] = picked, psc

    low = sorted(((score_of[ru], ru, assign[ru]) for ru in assign), key=lambda x: x[0])
    n_low = sum(1 for s, _, _ in low if s < 0.34)

    def walk(t, parent_ru=None):
        # sort each level ALPHABETICALLY by the English label (site sorts by match count;
        # we have none, so alphabetical) and drop a child that repeats its parent's name
        # (the «Belgium, Leuven…» under itself dup).
        items = []
        for ru, v in t.items():
            if ru == parent_ru:
                continue
            en = assign.get(ru, ru_to_en(ru))
            node = {"en": en}
            if v.get("children"):
                node["children"] = walk(v["children"], ru)
            items.append((en.lower(), ru, node))
        items.sort(key=lambda x: x[0])
        return {ru: node for _e, ru, node in items}

    out = walk(tree)
    _JSON.write_text(json.dumps(out, ensure_ascii=False, indent=2), "utf-8")
    dups = len(ru_names) - len(set(assign.values()))
    print(f"nodes: {len(ru_names)} | low-confidence (<0.34): {n_low} | "
          f"EN labels reused: {sum(1 for e in set(assign.values()) if list(assign.values()).count(e) > 1)}")


if __name__ == "__main__":
    main()
