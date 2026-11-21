"""6.5 (the owner, 20 Nov 2026 22:41: "Raise validity of the research tools to a much higher level. Make it extremely
valid."): what every research tool checks before a sentence may stand as sourced, and before a paper may be cited.

A CLAIM (a sentence of Research, of research writing or of a manual-writing statement, and the sentences of its papers
that back it) may not say more than its papers do:
  - a year it names is in what its papers say (or is a cited paper's own year);
  - a place it names (a country, a region, a city, or "Ghanaian", "Nigerian"…) is the papers' place;
  - "globally", "worldwide", "around the world" only when its papers speak of the world or of many countries;
  - "studies show", "research has found", "several studies" only when it cites two papers or more;
  - words in quotation marks are its papers' own words, letter for letter;
  - a cause ("causes", "leads to", "results in") never stands on papers that found only a link ("associated with");
  - a group of people it names (students, adolescents, children, women, nurses…) is a group its papers studied.
A SOURCE is checked with the registry that gave its DOI before it may be cited:
  - retracted papers are left out, whichever index found them (OpenAlex, which carries Retraction Watch's data);
  - the reference's data (authors, year, journal, volume, issue, pages) is the publisher's record at Crossref;
  - a DOI that doesn't exist, or whose record is a different paper, takes the paper out;
  - a paper with neither a DOI nor a named journal, publisher or institution can't be identified, and is left out.
Nothing here calls the network by itself: the server passes its own fetch function (the tests pass a made-up one)."""
import math
import re
import concurrent.futures as _futures

import stats66 as _st66

# ---------------------------------------------------------------- claims
_CITE_PAREN = re.compile(r"\((?:[^()]*?(?:1[5-9]|20)\d{2}[a-z]?|[^()]*?n\.\s?d\.|[^()]*?,\s*(?:in\s+press|forthcoming|no\s+date))[^()]*\)")   # (Mensah, 2021; Owusu, 2019)
_CITE_NUM = re.compile(r"\[\s*\d{1,3}(?:\s*[,;–\-]\s*\d{1,3})*\s*\]")                          # [1, 3]


_NARR = re.compile(r"(?:[A-Z\u00C0-\u024F][\w'\u2019\-]+\s+)?[A-Z\u00C0-\u024F][\w'\u2019\-]+(?:\s+(?:and|&)\s+[A-Z\u00C0-\u024F][\w'\u2019\-]+|,?\s+et al\.)?"
                   r"(?:['\u2019]s)?\s*\((?:1[5-9]|20)\d{2}[a-z]?(?:,\s*pp?\.\s*[\d\u2013\-]+)?\)")


def bare(sentence):
    """The sentence without its citations: author-date in brackets, numbers in square brackets, and a narrative
    citation's names with its year ("Paris (2019) found" is an author, not the city)."""
    s = _NARR.sub(" ", str(sentence or ""))
    s = _CITE_NUM.sub(" ", _CITE_PAREN.sub(" ", s))
    return re.sub(r"\s+", " ", s).strip()


def _norm(t):
    t = str(t or "").lower()
    t = t.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    t = t.replace("–", "-").replace("—", "-").replace(" ", " ")
    return re.sub(r"\s+", " ", t)


# ---- years
_YEAR = re.compile(r"(?<![\w.\-/])(1[89]\d{2}|20\d{2})(?![\w])")


_REPORTED = re.compile(r"\b(?:report\w*|found|finds?|show(?:s|ed|n)?|publish\w*|observ\w*|not(?:e|es|ed)|conclud\w*|argu\w*|"
                       r"describ\w*|examin\w*|stud(?:y|ies|ied)|survey(?:s|ed)?|conduct\w*|document\w*|estimat\w*|investigat\w*|"
                       r"explor\w*|assess\w*|reveal\w*|demonstrat\w*|indicat\w*|suggest\w*|wrote|writes|review\w*|analys\w*)\b", re.I)


def years_ok(sentence, evidence, pub_years=()):
    """Every year the sentence names (outside its citations) is in the evidence, or is a cited paper's own year. A paper's
    own year is the year it came out: said of its figures ("In 2021, 38% of students…") where the paper gives the year its
    data are from ("In 2019, 38%…"), it is the wrong year (6.6, round 23). "In 2021, Mensah reported…" stays."""
    have = set(_YEAR.findall(str(evidence or "")))
    pubs = {str(y) for y in pub_years if y}
    s = bare(sentence)
    for y in _YEAR.findall(s):
        if y in have:
            continue
        if y not in pubs:
            return False
        other = re.sub(r"(?<![\w.\-/])" + y + r"(?![\w])", " ", s)
        if have and not _REPORTED.search(s) and _FIGURE_SAID.search(other):
            return False
    return True


# ---- places (countries, regions, large cities; the words for their people)
_PLACE_NAMES = """
afghanistan|albania|algeria|angola|argentina|armenia|australia|austria|azerbaijan|bahamas|bahrain|bangladesh|barbados|belarus|
belgium|belize|benin|bhutan|bolivia|bosnia and herzegovina|botswana|brazil|brunei|bulgaria|burkina faso|burundi|cabo verde|
cape verde|cambodia|cameroon|canada|central african republic|chad|chile|china|colombia|comoros|costa rica|cote d'ivoire|
côte d'ivoire|ivory coast|croatia|cuba|cyprus|czechia|czech republic|denmark|djibouti|dominican republic|
democratic republic of the congo|republic of the congo|congo|ecuador|egypt|el salvador|equatorial guinea|eritrea|estonia|eswatini|
swaziland|ethiopia|fiji|finland|france|gabon|gambia|germany|ghana|greece|guatemala|guinea-bissau|guinea|guyana|haiti|
honduras|hungary|iceland|india|indonesia|iran|iraq|ireland|israel|italy|jamaica|japan|kazakhstan|kenya|kosovo|kuwait|
kyrgyzstan|laos|latvia|lebanon|lesotho|liberia|libya|lithuania|luxembourg|madagascar|malawi|malaysia|maldives|mali|malta|
mauritania|mauritius|mexico|moldova|mongolia|montenegro|morocco|mozambique|myanmar|namibia|nepal|netherlands|new zealand|
nicaragua|niger|nigeria|north korea|north macedonia|norway|oman|pakistan|palestine|panama|papua new guinea|paraguay|peru|
philippines|poland|portugal|qatar|romania|russia|rwanda|saudi arabia|senegal|serbia|seychelles|sierra leone|singapore|
slovakia|slovenia|somalia|south africa|south korea|korea|south sudan|spain|sri lanka|sudan|suriname|sweden|switzerland|syria|
taiwan|tajikistan|tanzania|thailand|timor-leste|togo|trinidad and tobago|tunisia|turkey|türkiye|turkmenistan|uganda|ukraine|
united arab emirates|united kingdom|great britain|england|scotland|wales|northern ireland|united states|uruguay|uzbekistan|
venezuela|vietnam|viet nam|yemen|zambia|zimbabwe|hong kong|puerto rico|
africa|sub-saharan africa|west africa|east africa|north africa|southern africa|central africa|asia|south asia|east asia|
southeast asia|south-east asia|central asia|europe|western europe|eastern europe|latin america|south america|north america|
central america|the caribbean|middle east|oceania|scandinavia|the sahel|
accra|kumasi|tamale|cape coast|takoradi|sekondi|koforidua|sunyani|bolgatanga|lagos|abuja|kano|ibadan|port harcourt|
enugu|nairobi|mombasa|kisumu|kampala|kigali|dar es salaam|dodoma|addis ababa|cairo|johannesburg|cape town|pretoria|durban|
lusaka|harare|maputo|luanda|kinshasa|dakar|abidjan|bamako|ouagadougou|niamey|lome|lomé|cotonou|freetown|monrovia|conakry|
banjul|rabat|casablanca|algiers|tunis|khartoum|mogadishu|lilongwe|blantyre|gaborone|windhoek|yaounde|yaoundé|douala|
london|paris|berlin|madrid|rome|beijing|shanghai|tokyo|seoul|delhi|new delhi|mumbai|karachi|lahore|dhaka|jakarta|manila|
bangkok|sydney|melbourne|toronto|new york|los angeles|mexico city|sao paulo|são paulo|rio de janeiro|buenos aires|tehran|
riyadh|dubai|istanbul""".replace("\n", "").split("|")
_PLACES = sorted({p.strip() for p in _PLACE_NAMES if p.strip()}, key=len, reverse=True)
# a few names are also ordinary words: they count only with a capital ("Ho", "Wa", "Niger", "Chad", "Guinea" stay names)
_CAPS_ONLY = {"chad", "niger", "guinea", "turkey", "china", "mali"}
# the words for a country's people and things ("Ghanaian students" is Ghana)
_DEMONYMS = {"ghanaian": "ghana", "nigerian": "nigeria", "kenyan": "kenya", "ugandan": "uganda", "tanzanian": "tanzania",
             "ethiopian": "ethiopia", "rwandan": "rwanda", "south african": "south africa", "egyptian": "egypt",
             "cameroonian": "cameroon", "senegalese": "senegal", "malawian": "malawi", "zambian": "zambia",
             "zimbabwean": "zimbabwe", "togolese": "togo", "beninese": "benin", "ivorian": "cote d'ivoire",
             "burkinabe": "burkina faso", "malian": "mali", "sudanese": "sudan", "moroccan": "morocco",
             "algerian": "algeria", "tunisian": "tunisia", "liberian": "liberia", "sierra leonean": "sierra leone",
             "gambian": "gambia", "botswanan": "botswana", "namibian": "namibia", "mozambican": "mozambique",
             "angolan": "angola", "congolese": "congo", "american": "united states", "british": "united kingdom",
             "scottish": "scotland", "irish": "ireland", "canadian": "canada", "mexican": "mexico", "brazilian": "brazil",
             "argentine": "argentina", "chilean": "chile", "colombian": "colombia", "peruvian": "peru", "indian": "india",
             "pakistani": "pakistan", "bangladeshi": "bangladesh", "nepalese": "nepal", "sri lankan": "sri lanka",
             "malaysian": "malaysia", "iranian": "iran", "iraqi": "iraq", "saudi": "saudi arabia", "israeli": "israel",
             "palestinian": "palestine", "australian": "australia", "belgian": "belgium", "swiss": "switzerland",
             "austrian": "austria", "african": "africa", "asian": "asia", "european": "europe", "latin american": "latin america",
             "caribbean": "the caribbean", "middle eastern": "middle east", "sub-saharan": "sub-saharan africa"}
# (the words that are also a language's name, "English", "French", "Chinese", "Swahili"…, are never read as a place:
# "taught in English" says nothing of England)
# a place and its other names (any of them in the evidence is the same place)
_SAME_PLACE = {"united states": ("united states", "usa", "u.s.", "us", "america", "american"),
               "united kingdom": ("united kingdom", "uk", "u.k.", "britain", "british", "england", "scotland", "wales"),
               "cote d'ivoire": ("cote d'ivoire", "côte d'ivoire", "ivory coast", "ivorian"),
               "côte d'ivoire": ("cote d'ivoire", "côte d'ivoire", "ivory coast", "ivorian"),
               "ivory coast": ("cote d'ivoire", "côte d'ivoire", "ivory coast", "ivorian"),
               "turkey": ("turkey", "türkiye", "turkish"), "türkiye": ("turkey", "türkiye", "turkish"),
               "vietnam": ("vietnam", "viet nam", "vietnamese"), "viet nam": ("vietnam", "viet nam", "vietnamese"),
               "eswatini": ("eswatini", "swaziland"), "swaziland": ("eswatini", "swaziland"),
               "lome": ("lome", "lomé"), "lomé": ("lome", "lomé"), "yaounde": ("yaounde", "yaoundé"), "yaoundé": ("yaounde", "yaoundé"),
               "sao paulo": ("sao paulo", "são paulo"), "são paulo": ("sao paulo", "são paulo"),
               "south-east asia": ("south-east asia", "southeast asia"), "southeast asia": ("south-east asia", "southeast asia"),
               "korea": ("korea", "south korea", "korean"), "congo": ("congo", "democratic republic of the congo", "drc",
                                                                       "republic of the congo", "congolese")}
_PLACE_RX = re.compile(r"(?<![\w\-])(" + "|".join(re.escape(p) for p in _PLACES) + r")(?![\w\-])", re.I)
_DEMONYM_RX = re.compile(r"(?<![\w\-])(" + "|".join(re.escape(d) for d in sorted(_DEMONYMS, key=len, reverse=True)) + r")(?![\w\-])", re.I)
_ABBR_RX = re.compile(r"(?<![\w.])(USA|U\.S\.|US|UK|U\.K\.|UAE|DRC)(?![\w$])")         # in capitals only ("us" is a word; "US$" is money)
_ABBR = {"usa": "united states", "u.s.": "united states", "us": "united states", "uk": "united kingdom",
         "u.k.": "united kingdom", "uae": "united arab emirates", "drc": "congo"}
# words that, before a place name, make it part of a name, not where something happened ("Lagos State University" is
# still Lagos; "New Mexico" is not Mexico: the longer names come first)


def places(text):
    """The places a text names (lower case, each once): countries, regions, large cities, and those named by the words
    for their people ("Ghanaian" is Ghana)."""
    t = str(text or "")
    out = []
    for m in _PLACE_RX.finditer(t):
        name = m.group(1).lower()
        if name in _CAPS_ONLY and not m.group(1)[:1].isupper():
            continue
        out.append(name)
    for m in _DEMONYM_RX.finditer(t):
        word = m.group(1).lower()
        if word in ("american", "african", "asian", "european") and re.match(r"\s+(?:psychological|medical|journal|association|society)\b",
                                                                              t[m.end():m.end() + 30], re.I):
            continue                    # a society's name ("American Psychological Association")
        out.append(_DEMONYMS[word])
    for m in _ABBR_RX.finditer(t):
        out.append(_ABBR[m.group(1).lower()])
    seen = []
    for p in out:
        if p not in seen:
            seen.append(p)
    return seen


def _place_in(place, evidence_places, evidence_text):
    names = _SAME_PLACE.get(place, (place,))
    if any(n in evidence_places for n in names):
        return True
    low = _norm(evidence_text)
    return any(re.search(r"(?<![\w\-])" + re.escape(n) + r"(?![\w\-])", low) for n in names if len(n) > 3)


def places_ok(sentence, evidence):
    """Every place the sentence names is a place its evidence names (the same place by another of its names counts:
    "Ghanaian" and "Ghana"; "the UK" and "Britain"). A place inside the evidence's own words is enough; a region is not
    the countries in it, nor a country the region ("in Africa" from a study in Ghana is more than the study says)."""
    ev_places = places(evidence)
    return all(_place_in(p, ev_places, evidence) for p in places(bare(sentence)))


# ---- the world
_WORLD_CLAIM = re.compile(r"\b(?:global(?:ly)?|world-?wide|around the (?:world|globe)|across the (?:world|globe)|in the world|"
                          r"the world over|internationally|across (?:many|most|all) countries|in (?:many|most) countries|"
                          r"the world's)\b", re.I)
_WORLD_EVID = re.compile(r"\b(?:global(?:ly)?|world-?wide|world|international(?:ly)?|(?:\d+|many|several|multiple|all|low-|middle-|"
                         r"high-|\w+-income)\s*countries|nations|multi-?country|cross-?national|across countries)\b", re.I)


def world_ok(sentence, evidence):
    """"Globally", "worldwide", "around the world" only when the evidence speaks of the world or of many countries."""
    return not _WORLD_CLAIM.search(bare(sentence)) or bool(_WORLD_EVID.search(str(evidence or "")))


# ---- more than one study
_MANY_CLAIM = re.compile(r"\b(?:(?:several|many|numerous|multiple|various|most|other|previous|earlier|recent|existing|prior)\s+"
                         r"(?:studies|researchers|authors|reports|trials|surveys|scholars|reviews)|"
                         r"studies\s+(?:have\s+|also\s+|consistently\s+)*(?:show|shown|showed|suggest|suggested|found|find|reported|"
                         r"report|indicate|indicated|demonstrate|demonstrated|confirm|confirmed|agree|revealed|reveal|documented|noted)|"
                         r"research(?:ers)?\s+(?:has\s+|have\s+|also\s+|consistently\s+)*(?:shows?|shown|showed|suggests?|suggested|found|"
                         r"finds|reported|reports|indicates?|indicated|demonstrated?|demonstrates|confirmed|confirms|revealed|reveals)|"
                         r"(?:the\s+)?(?:literature|evidence base|body of (?:research|evidence|literature)|growing body)\s+"
                         r"(?:\w+\s+){0,3}?(?:shows?|suggests?|indicates?|is consistent|agrees?|points)|"
                         r"(?:it\s+is|it\s+has\s+been)\s+(?:well[\s-]+)?(?:established|widely\s+(?:known|accepted|reported|recognised|"
                         r"recognized|documented|acknowledged))|a growing body of|a large body of|a body of (?:research|evidence|literature)|"
                         r"consistently\s+(?:shown|found|reported))\b", re.I)


# a paper that is itself about many studies (a review, a meta-analysis, a pooled analysis of 24 studies) may back them
_MANY_EVID = re.compile(r"\b(?:review(?:ed|s)?|meta-?analys[ie]s|pooled|(?:\d+|two|three|four|five|six|seven|eight|nine|ten|several|many|"
                        r"multiple|numerous|previous|prior|earlier|other|included)\s+(?:[\w-]+\s+){0,2}?(?:studies|trials|papers|articles|"
                        r"surveys|reports))\b", re.I)


# (6.6, round 32) "several", "many", "numerous", "various" or "most" studies, "a growing (large) body of", "it is well
# established": more than two; two papers are "two studies" or "studies"
_MANY3 = re.compile(r"\b(?:several|many|numerous|various|most)\s+(?:studies|researchers|authors|reports|trials|surveys|scholars|reviews)\b|"
                    r"\ba\s+(?:growing|large|substantial|considerable|vast)\s+body\s+of\b|\b(?:it\s+is|it\s+has\s+been)\s+(?:well[\s-]+)?"
                    r"(?:established|widely\s+(?:known|accepted|reported|recognised|recognized|documented|acknowledged))\b"
                    # 6.6, round 60: "it is (well) known that", "is (widely) known to"
                    r"|\bit\s+(?:is|has\s+(?:long\s+)?been)\s+(?:well[\s-]+|widely\s+|long\s+|now\s+)?known\s+that\b"
                    r"|\b(?:is|are)\s+(?:well[\s-]+|widely\s+)?known\s+to\s+(?!be\s+(?:called|named))", re.I)


def many_ok(sentence, n_papers, evidence=""):
    """"Studies show", "research has found": two cited papers at least; "several", "many", "most" studies, "a growing body
    of research", "it is well established": three at least (6.6, round 32); or one paper that is itself about many studies
    (a review, a meta-analysis: "A review of 24 studies found…")."""
    s = bare(sentence)
    if bool(_MANY_EVID.search(str(evidence or ""))):
        return True
    if _MANY3.search(s):
        return (n_papers or 0) >= 3
    return (n_papers or 0) >= 2 or not _MANY_CLAIM.search(s)


# ---- how strong the evidence is (6.6, round 60)
_STRONG_EV = re.compile(r"\b(?:strong|stronger|robust|compelling|overwhelming|solid|substantial|convincing|clear|firm|ample|extensive|"
                        r"conclusive|definitive|unequivocal|irrefutable|high[\s-]quality)\s+(?:\w+\s+)?evidence\b"
                        r"|\bevidence\s+(?:is|was|remains|has\s+been|appears)\s+(?:\w+\s+)?(?:strong|robust|compelling|overwhelming|clear|"
                        r"convincing|unequivocal|firm|solid|conclusive)\b"
                        r"|\b(?:firmly|clearly|well|convincingly)\s+(?:established|documented|demonstrated|proven|shown)\b", re.I)


def evidence_ok(sentence, quoted, n_papers=1):
    """"Strong evidence", "robust evidence", "compelling", "the evidence is clear", "firmly established", "well
    documented": how strong the evidence is, which one survey can't say of itself. Only where the papers' own sentences
    speak of their evidence so (a review that rates its evidence strong), or the sentence rests on three papers or more
    (as "it is well established" does, round 32)."""
    s = bare(sentence)
    if not _STRONG_EV.search(s):
        return True
    return bool(_STRONG_EV.search(str(quoted or ""))) or (n_papers or 0) >= 3


# ---- words in quotation marks
_QUOTED = re.compile(r"[“\"]([^“”\"]{3,400})[”\"]|(?<![\w])‘([^‘’]{3,400})’(?![\w])")


def _letters(t):
    return re.sub(r"[^\w%]+", " ", _norm(t)).strip()


def quotes_ok(sentence, papers_text):
    """Words the sentence puts in quotation marks (three or more) are the papers' own words, in the same order."""
    have = " " + _letters(papers_text) + " "
    for m in _QUOTED.finditer(str(sentence or "")):
        q = m.group(1) or m.group(2) or ""
        if len(re.findall(r"\w+", q)) < 3:
            continue
        if " " + _letters(q) + " " not in have:
            return False
    return True


# ---- a cause where the papers found a link
_CAUSE_CLAIM = re.compile(r"\b(?:caus(?:e|es|ed|ing)|leads?\s+to|led\s+to|leading\s+to|results?\s+in|resulted\s+in|resulting\s+in|"
                          r"(?:is|are|was|were)\s+responsible\s+for|is\s+the\s+(?:main\s+|primary\s+|leading\s+)?(?:cause|reason)|"
                          r"brings?\s+about|brought\s+about|gives?\s+rise\s+to|gave\s+rise\s+to)\b", re.I)
_LINK_EVID = re.compile(r"\b(?:associat\w*|correlat\w*|linked|links?\b|relationship|related\s+to|relation\s+to|predict\w*|"
                        r"odds|risk\s+factors?)\b", re.I)
_CAUSE_EVID = re.compile(r"\b(?:caus\w*|effects?\b|effective\w*|impact\w*|led\s+to|leads?\s+to|leading\s+to|result\w*\s+in|due\s+to|"
                         r"because|randomi[sz]\w*|trials?\b|experiment\w*|intervention\w*|attribut\w*|contribut\w*|"
                         r"responsible\s+for|determinants?)\b", re.I)


def cause_ok(sentence, quoted):
    """A sentence that says one thing causes another stands only when its papers' sentences speak of a cause, an
    effect or a trial; papers that found a link ("associated with", "correlated") back a link, not a cause."""
    if not _CAUSE_CLAIM.search(bare(sentence)):
        return True
    q = str(quoted or "")
    return bool(_CAUSE_EVID.search(q)) or not _LINK_EVID.search(q)


# ---- the people studied
_GROUPS = {
    "students": r"students?|pupils?|learners?|undergraduates?|schoolchildren|school\s+children|schoolgirls?|schoolboys?",
    "adolescents": r"adolescen\w*|teen(?:ager)?s?|youths?|young\s+people|young\s+persons?",
    "children": r"child(?:ren)?|kids|infants?|toddlers?|pre-?school\w*|under-?fives?|newborns?|neonates?|pupils?|schoolchildren|"
                r"school\s+children|schoolgirls?|schoolboys?|(?:primary|basic|elementary)\s+school\w*",
    "older adults": r"elderly|older\s+(?:adults?|people|persons?|women|men)|aged\s+(?:6[0-9]|7[0-9])\s*(?:years|and\s+(?:over|above))",
    "women": r"women|woman|females?|mothers?|girls?|schoolgirls?|pregnan\w*|maternal|antenatal",
    "men": r"men|man|males?|fathers?|boys?|schoolboys?",
    "nurses": r"nurses?|nursing\s+staff|midwi(?:fe|ves)",
    "doctors": r"doctors?|physicians?|clinicians?|surgeons?|medical\s+officers?",
    "health workers": r"health(?:care)?\s+(?:care\s+)?(?:workers?|professionals?|providers?|personnel|staff)",
    "patients": r"patients?|inpatients?|outpatients?",
    "teachers": r"teachers?|educators?|lecturers?|instructors?|tutors?",
    "farmers": r"farmers?|smallholders?",
}
_GROUP_RX = {g: re.compile(r"(?<![\w\-])(?:" + rx + r")(?![\w\-])", re.I) for g, rx in _GROUPS.items()}
# a health worker is a nurse or a doctor too
_GROUP_ALSO = {"health workers": ("nurses", "doctors"), "students": ("adolescents",), "adolescents": ("students",)}


def groups(text):
    t = str(text or "")
    return [g for g, rx in _GROUP_RX.items() if rx.search(t)]


def groups_ok(sentence, papers_text):
    """Every group of people the sentence names is one its papers studied (a health worker may be a nurse or a doctor
    there; secondary-school students and adolescents are the same young people)."""
    have = set(groups(papers_text))
    for g in groups(bare(sentence)):
        if g in have or any(x in have for x in _GROUP_ALSO.get(g, ())):
            continue
        return False
    return True


# ---- 6.6, round 36: who within a group: the school level, the year of study, rural or urban, public or private
_QUAL = (
    ({"primary": r"\b(?:primary|elementary)[\s-]+(?:schools?|pupils?|students?|education|level|class(?:es)?)\b",
      "basic": r"\bbasic[\s-]+schools?\b|\bbasic[\s-]+education\b",          # (6.6, round 42: Ghana's basic school is primary and JHS)
      "jhs": r"\bjunior[\s-]+(?:high|secondary)\b|\bJHS\b|\bmiddle[\s-]+schools?\b",
      "shs": r"\bsenior[\s-]+(?:high|secondary)\b|\bSHS\b",
      "secondary": r"(?<!junior\s)(?<!senior\s)(?<!junior-)(?<!senior-)\bsecondary[\s-]+(?:schools?|students?|pupils?|education|level)\b|"
                   r"(?<!junior\s)(?<!senior\s)(?<!junior-)(?<!senior-)\bhigh[\s-]+schools?\b",
      "tertiary": r"\btertiary\b|\buniversit(?:y|ies)\b|\bcolleges?\b|\bundergraduates?\b|\bpolytechnics?\b",
      "postgraduate": r"\bpost-?graduates?\b|\bgraduate\s+students?\b|\bmaster['’]?s\s+students?\b|\bdoctoral\s+students?\b|"
                      r"\bPh\.?D\.?\s+students?\b"},
     {"primary": {"primary", "basic"}, "basic": {"basic", "primary", "jhs"}, "jhs": {"jhs", "secondary", "basic"}, "shs": {"shs", "secondary"},
      "secondary": {"secondary", "jhs", "shs"},
      "tertiary": {"tertiary", "postgraduate"}, "postgraduate": {"postgraduate"}}),
    ({"y1": r"\b(?:first|1st)[\s-]+years?\s+(?:\w+\s+)?(?:students?|undergraduates?|nurses?|pupils?|learners?|trainees?)\b|\bfreshm[ae]n\b|"
            r"\blevel[\s-]+100\b",
      "y2": r"\b(?:second|2nd)[\s-]+years?\s+(?:\w+\s+)?(?:students?|undergraduates?|nurses?|pupils?|learners?|trainees?)\b|\bsophomores?\b|"
            r"\blevel[\s-]+200\b",
      "y3": r"\b(?:third|3rd)[\s-]+years?\s+(?:\w+\s+)?(?:students?|undergraduates?|nurses?|pupils?|learners?|trainees?)\b|\blevel[\s-]+300\b",
      "y4": r"\b(?:fourth|4th|final|last)[\s-]+years?\s+(?:\w+\s+)?(?:students?|undergraduates?|nurses?|pupils?|learners?|trainees?)\b|"
            r"\blevel[\s-]+400\b"},
     {"y1": {"y1"}, "y2": {"y2"}, "y3": {"y3"}, "y4": {"y4"}}),
    ({"rural": r"\brural\b", "urban": r"(?<!-)\burban\b", "periurban": r"\bperi-?urban\b"},
     {"rural": {"rural"}, "urban": {"urban", "periurban"}, "periurban": {"periurban", "urban"}}),
    ({"public": r"\b(?:public|government|state)[\s-]+(?:schools?|universit(?:y|ies)|hospitals?|health\s+facilit(?:y|ies)|clinics?|sector)\b",
      "private": r"\bprivate[\s-]+(?:schools?|universit(?:y|ies)|hospitals?|health\s+facilit(?:y|ies)|clinics?|sector)\b"},
     {"public": {"public"}, "private": {"private"}}),
)
_QUAL_RX = tuple(({k: re.compile(v, re.I) for k, v in fam.items()}, compat) for fam, compat in _QUAL)


def qualifiers_ok(sentence, evidence):
    """Who within a group (6.6, round 36): secondary-school students are not university students, first-year students are
    not final-year students, rural people are not urban people, private schools are not public schools. Said only where
    the papers name another of the same kind and never this one (a paper that names none is not judged)."""
    s, ev = bare(sentence), str(evidence or "")
    for fam, compat in _QUAL_RX:
        have = {k for k, rx in fam.items() if rx.search(ev)}
        if not have:
            continue
        for k, rx in fam.items():
            if rx.search(s) and not (compat[k] & have):
                return False
    return True


# ---------------------------------------------------------------- 6.5, further (the owner, 20 Nov 2026 23:50: "Increase the
# validity of the research tools further"): more of what a sentence says must be its papers' own
# ---- the kind of study ("a randomised trial found…" from a survey)
_DESIGNS = {
    "trial": (r"randomi[sz]ed|randomly\s+(?:assigned|allocated)|(?:controlled|clinical|field|cluster)\s+trials?|\bRCTs?\b|\btrials?\b",
              r"randomi[sz]\w*|randomly\s+(?:assigned|allocated)|\btrials?\b|\bRCTs?\b"),
    "meta-analysis": (r"meta-?analys[ie]s", r"meta-?analys[ie]s"),
    "systematic review": (r"systematic(?:\s+literature)?\s+reviews?", r"systematic(?:\s+literature)?\s+reviews?|meta-?analys[ie]s"),
    # ("this review", "the present review" are the piece being written, not the cited paper: "In this review, … (Mensah, 2021)")
    "review": (r"(?<!\bthis\s)(?<!\bpresent\s)(?<!\bcurrent\s)(?<!\bour\s)(?:scoping|narrative|literature|integrative|umbrella)\s+reviews?"
               r"|\ba\s+review\b|\bthe\s+review\b",
               r"\breview\w*|meta-?analys[ie]s"),
    "cohort": (r"\bcohort\b|\blongitudinal\b|\bprospective(?:ly)?\b|\bfollowed\s+(?:up|for)\b",
               r"\bcohort\b|\blongitudinal\b|\bprospective\w*|\bfollow(?:ed)?[\s-]+(?:up|for)\b|\bover\s+(?:\w+\s+){0,2}(?:years|months)\b"),
    "cross-sectional": (r"cross-?sectional", r"cross-?sectional|\bsurvey\w*|\bquestionnaires?\b"),
    "case-control": (r"case-?control", r"case-?control"),
    "qualitative": (r"\bqualitative\b|\bfocus\s+groups?\b|\bin-?depth\s+interviews?\b|\bthematic(?:ally)?\b",
                    r"\bqualitative\b|\binterview\w*|\bfocus\s+groups?\b|\bthematic\w*"),
    "experiment": (r"quasi-?experiment\w*|\bexperiment(?:al|s)?\b", r"experiment\w*|randomi[sz]\w*|\btrials?\b|\bintervention\w*"),
}
_DESIGN_RX = {k: (re.compile(c, re.I), re.compile(e, re.I)) for k, (c, e) in _DESIGNS.items()}


def design_ok(sentence, evidence):
    """The kind of study a sentence names (a randomised trial, a cohort, a meta-analysis, a qualitative study…) is the
    kind its papers are: a survey is not a trial, one survey is not a longitudinal study."""
    s, ev = bare(sentence), str(evidence or "")
    return all(evid.search(ev) for k, (claim, evid) in _DESIGN_RX.items() if claim.search(s))


# ---- "significant", "proves", "always", "the most…": words that claim more than a paper may
_SIGNIF = re.compile(r"\bsignifican(?:t|tly|ce)\b", re.I)
_SIGNIF_EVID = re.compile(r"\bsignifican\w*|\bp\s*[<=≤>]\s*0?\.\d|\bp-?values?\b|confidence\s+intervals?|\b\d+\s*%\s*CI\b|\bCI\b|"
                          r"odds\s+ratios?|\b(?:a?OR|RR|HR|IRR)\s*[=:]?\s*\d|\bstatistically\b|\bχ²|\bchi-?square", re.I)
_NEGATED = re.compile(r"\b(?:no|not|nor|neither|without|never|n't|failed\s+to|did\s+not|does\s+not|was\s+not|were\s+not)\b", re.I)


_SIG_WORD = re.compile(r"\b(?:non-?|in)?significan(?:t|tly|ce)\b", re.I)
# the words that turn it to "not": "no significant", "not (statistically) significant", "did not differ significantly",
# "wasn't significant", "did not reach significance", "failed to reach statistical significance", "lack of significance",
# "without significant"; never "not only significant", nor a "no" further back ("those with no phone slept significantly more")
_SIG_NEG = re.compile(r"(?:\b(?:no|not|never|nor|neither)\s+(?!only\b|just\b|merely\b)(?:\w+\s+)?|n['’]t\s+(?:\w+\s+)?|"
                      r"\b(?:fail(?:s|ed)?\s+to|did\s+not|does\s+not|do\s+not|didn['’]t|doesn['’]t)\s+(?:reach|achieve|attain|show|"
                      r"demonstrate|find|yield|meet)\s+(?:\w+\s+)?|\b(?:lack|absence)\s+of\s+|\bwithout\s+)"
                      r"(?:statistical(?:ly)?\s+)?(?=significan)", re.I)
_P_SAID = re.compile(r"\bp\s*(<=|>=|<|=|≤|>|≥)\s*(0?\.\d+|1(?:\.0+)?)\b", re.I)
_NS_SAID = re.compile(r"\(\s*n\.?\s?s\.?\s*\)|\bn\.s\.", re.I)


def _sig_said(text, with_p=False):
    """(significant, not significant) as the text says them: "significantly less", "a significant difference";
    "not significant", "no significant", "non-significant", "did not reach significance". With with_p, a p below .05 is
    significant and one above .05 not (p = .05 itself says neither)."""
    t = str(text or "")
    pos = neg = False
    nots = {m.end() for m in _SIG_NEG.finditer(t)}
    for m in _SIG_WORD.finditer(t):
        if m.group(0).lower().startswith(("non", "insig")) or m.start() in nots:
            neg = True
        else:
            pos = True
    if with_p:
        for op, v in _P_SAID.findall(t):
            try:
                x = float(v)
            except ValueError:
                continue
            if op in ("<", "≤", "<=") and x <= 0.05:
                pos = True
            elif op in ("=",) and x < 0.05:
                pos = True
            elif (op in (">", "≥", ">=") and x >= 0.05) or (op == "=" and x > 0.05):
                neg = True
        if _NS_SAID.search(t):
            neg = True
    return pos, neg


def significance_ok(sentence, quoted):
    """"Significant(ly)" only where the papers' sentences speak of significance or give the statistics (a p-value, a
    confidence interval, an odds ratio): in academic writing it is a statistical claim. And it keeps its side (6.6, round
    23): "not significant" where the papers found it significant (and never not), or "significantly" where they found no
    significant result (and never one), turns the finding over."""
    s = bare(sentence)
    if not _SIG_WORD.search(s):
        return True
    q = str(quoted or "")
    if not (_SIGNIF_EVID.search(q) or re.search(r"insignifican|\(\s*n\.?\s?s\.?\s*\)", q, re.I)):
        return False
    s_pos, s_neg = _sig_said(s)
    q_pos, q_neg = _sig_said(q, with_p=True)
    if s_neg and q_pos and not q_neg:
        return False
    if s_pos and q_neg and not q_pos:
        return False
    return True


_CERTAIN = re.compile(r"\b(prov(?:e|es|ed|en|ing)|conclusive(?:ly)?|definitive(?:ly)?|undeniabl[ey]|indisputabl[ey]|unquestionabl[ey]|"
                      r"undoubted(?:ly)?|certainly|beyond\s+(?:any\s+)?doubt|always|never|without\s+exception"
                      # 6.6, round 60: "definitely", "clearly shows", "without (a) doubt"
                      r"|definitely|clearly(?=\s+(?:show|demonstrat|prov|establish|indicat|confirm))|without\s+(?:a\s+|any\s+)?doubt)\b", re.I)


def _root(w):
    w = re.sub(r"\s+", " ", w.lower())
    if w.startswith("beyond") or re.match(r"without (?:a |any )?doubt", w):
        return "beyond"                    # "beyond doubt" and "without doubt" are the same claim
    return w[:5]


def certainty_ok(sentence, quoted):
    """"Proves", "conclusively", "definitively", "always", "never"…: only when the papers' own sentences say as much.
    6.6, round 111: "never" of a share ("38% had never attended a training") or of an experience ("had never used") is the
    papers' "not" ("had not attended any"), not a claim that something never happens."""
    q = {_root(m.group(1)) for m in _CERTAIN.finditer(str(quoted or ""))}
    s = bare(sentence)
    for m in _CERTAIN.finditer(s):
        r = _root(m.group(1))
        if r in q:
            continue
        if r == "never":
            before = re.split(r"[;:]|\.(?!\d)|,\s*(?:but|while|whereas)\b", s[:m.start()])[-1]
            if (re.search(r"\d+(?:\.\d+)?\s?%|\b\d+\s+(?:of|out\s+of)\s+\d+", before) or re.search(r"\b(?:had|have|has)\s*$", before)) \
                    and re.search(r"\b(?:not|no|none|without)\b|n['’]t\b", str(quoted or ""), re.I):
                continue
        return False
    return True


_SUPERL = re.compile(r"\bthe\s+(most|least)\s+[a-z][\w-]+|\bthe\s+(best|worst|largest|smallest|highest|lowest|greatest|strongest|weakest|"
                     r"biggest|leading|only|commonest)\b|\b(most|least)\s+(?:common(?:ly)?|frequent(?:ly)?|effective|important|"
                     r"prevalent|significant|reported|used)\b", re.I)


def superlative_ok(sentence, quoted):
    """"The most common", "the leading cause", "the only study": only when the papers' sentences say the same."""
    q = str(quoted or "").lower()
    for m in _SUPERL.finditer(bare(sentence)):
        w = (m.group(1) or m.group(2) or m.group(3) or "").lower()
        if w == "commonest":
            w = "common"
        if not re.search(r"\b" + re.escape(w) + r"\b", q) and not (w in ("most", "commonest") and re.search(r"\bcommonest\b", q)):
            return False
    return True


# ---- "most students…": a share the papers' figures don't give
_MOST = re.compile(r"\b(most|the\s+majority|a\s+majority|nearly\s+all|almost\s+all|virtually\s+all|the\s+vast\s+majority|a\s+minority|"
                   r"few|fewer\s+than\s+half|more\s+than\s+half|over\s+half|about\s+half|half)\s+(?:of\s+)?(?:the\s+|these\s+|those\s+)?"
                   r"(?:[\w-]+\s+){0,3}?(?:students|pupils|participants|respondents|children|adolescents|women|men|nurses|patients|teachers|"
                   r"adults|people|workers|farmers|mothers|households|schools|cases|users|learners|subjects)\b", re.I)
_SHARE_WORDS = {"most": (50.01, 100), "the majority": (50.01, 100), "a majority": (50.01, 100), "nearly all": (85, 100),
                "almost all": (85, 100), "virtually all": (90, 100), "the vast majority": (75, 100), "a minority": (0, 49.99),
                "few": (0, 30), "fewer than half": (0, 49.99), "more than half": (50.01, 100), "over half": (50.01, 100),
                "about half": (40, 60), "half": (45, 55)}


def shares_ok(sentence, quoted):
    """"Most students", "the majority", "nearly all", "few": the papers' sentences say the same, or give a share that is
    that (most: more than half; nearly all: 85% or more; few: under 30%)."""
    q = str(quoted or "")
    pcts = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", q)]
    for m in _MOST.finditer(bare(sentence)):
        w = re.sub(r"\s+", " ", m.group(1).lower())
        if re.search(r"\b" + re.escape(w) + r"\b", q, re.I):
            continue
        if w in ("most", "the majority", "a majority") and re.search(r"\b(?:most|majority)\b", q, re.I):
            continue
        lo, hi = _SHARE_WORDS.get(w, (None, None))
        if lo is not None and any(lo <= v <= hi for v in pcts):
            continue
        return False
    return True


# ---- the direction of a finding ("reduced anxiety" where the paper found anxiety rose)
_UP = r"increas\w*|higher|rose|rise[sn]?|rising|greater|grew|grow(?:n|s|ing)?|gain(?:s|ed|ing)?|rais(?:e|es|ed|ing)|elevat\w*|more"
_DOWN = (r"decreas\w*|lower(?:s|ed|ing)?|fell|fall(?:s|en|ing)?|declin\w*|fewer|reduc\w*|drop(?:s|ped|ping)?|shrank|shrunk|"
         r"diminish\w*|less")
_STOPW = {"the", "a", "an", "of", "in", "on", "for", "and", "or", "to", "their", "its", "his", "her", "than", "with", "by", "among", "at",
          "levels", "level", "rates", "rate", "scores", "score", "risk", "odds", "likely", "use", "number", "amount", "this", "that",
          "these", "those", "was", "were", "is", "are", "be", "been", "after", "before", "during", "students", "participants"}


# 6.6, round 23: better and worse ("sleep quality worsened" where it improved), earlier and later ("fell asleep earlier"
# where they fell asleep later)
_BETTER = r"improv(?:e|es|ed|ing|ement|ements)|better"
_WORSE = r"wors(?:e|en|ens|ened|ening)|deteriorat\w*|poorer"
_DIR_WORDS = (("up", _UP), ("down", _DOWN), ("better", _BETTER), ("worse", _WORSE), ("earlier", r"earlier"), ("later", r"later"))
_DIR_OTHER = {"up": "down", "down": "up", "better": "worse", "worse": "better", "earlier": "later", "later": "earlier"}
_DIR_ANY = "|".join(rx for _d, rx in _DIR_WORDS)


def _dir_pairs(text):
    """(direction, what) pairs: "reduced anxiety", "anxiety decreased", "lower sleep quality", "sleep improved",
    "fell asleep later"."""
    out = set()
    t = re.sub(r"[^\w\s%-]", " ", str(text or "").lower())
    words = t.split()
    for i, w in enumerate(words):
        d = next((k for k, rx in _DIR_WORDS if re.fullmatch(rx, w)), None)
        if not d:
            continue
        if any(re.fullmatch(r"not|no|nor|never|without|neither|n't", x) for x in words[max(0, i - 3):i]):
            continue
        if "than" in words[i + 1:i + 7]:
            # a comparison: "girls slept less than boys" is "boys slept more than girls"; but against the same side ("…than at
            # weekends"), "more" for the paper's "less" turns it over (6.6, round 26)
            k = words.index("than", i + 1)
            other = [x for x in words[k + 1:k + 6] if x not in _STOPW and len(x) >= 3 and not re.fullmatch(_DIR_ANY, x)]
            if other:
                out.add((d, "than:" + "_".join(x[:6] for x in other[:2])))     # ("than in the wet season": "wet season")
            continue
        after = [x for x in words[i + 1:i + 4] if x not in _STOPW and len(x) > 3 and not re.fullmatch(_DIR_ANY, x)]
        before = [x for x in words[max(0, i - 3):i] if x not in _STOPW and len(x) > 3 and not re.fullmatch(_DIR_ANY, x)]
        for x in (after[:1] + before[-1:]):
            out.add((d, x[:6]))
    return out


def direction_ok(sentence, quoted):
    """A finding stated in the direction its papers found: when the sentence says something rose (or fell, improved,
    worsened, came earlier or later) and the papers' sentences say that same thing went the other way, and never this
    way, the sentence turned the finding over."""
    s, q = _dir_pairs(bare(sentence)), _dir_pairs(quoted)
    for d, x in s:
        if (_DIR_OTHER[d], x) in q and (d, x) not in q:
            return False
    return True


# ---------------------------------------------------------------- 6.5, higher (the owner, 21 Nov 2026 00:16: "Raise it
# further higher"): trends, scale, shares in words, more words of cause, hedges, "recent", "the first", a finding's polarity
# ---- a trend ("is rising", "has increased", "increasingly") only where the papers' sentences speak of a change over time
_TREND_NOUNS = (r"(?:trends?|prevalence|incidence|rates?|numbers?|burden|use|usage|concerns?|problems?|popularity|interest|cases|"
                r"levels?|demand|reliance|dependence|uptake|adoption|ownership|access|exposure)")
_TREND_MORE = (r"|\b(?:keeps?|kept|keeping|continues?|continued|continuing)\s+(?:to\s+)?(?:rise|rising|grow|growing|increase|"
               r"increasing|decline|declining|decrease|decreasing|fall|falling|climb|climbing|drop|dropping|spread|spreading)\b"
               r"|\bmore\s+and\s+more\b|\bfewer\s+and\s+fewer\b|\bever[-\s](?:increasing|growing|rising|higher)\b"
               r"|\b(?:is|are|has\s+been|have\s+been)\s+becoming\s+(?:more|less|increasingly)\b"
               r"|\b(?:has|have)\s+become\s+(?:more|less|increasingly)\s+(?:common|prevalent|popular|widespread|frequent|rare|"
               r"usual|accepted|available)\b")
_TREND_CLAIM = re.compile(
    r"\b(?:is|are|has\s+been|have\s+been|was|were)\s+(?:(?:steadily|rapidly|sharply|gradually|steeply|fast|quickly|consistently|"
    r"continuously|dramatically|still|now)\s+)?(?:rising|increasing|growing|declining|decreasing|falling|dropping|climbing|"
    r"escalating|surging|soaring|on\s+the\s+(?:rise|increase))\b"
    r"|\b(?:has|have)\s+(?:(?:steadily|rapidly|sharply|gradually|steeply|dramatically|significantly|substantially|markedly|"
    r"considerably|greatly|almost|nearly|more\s+than)\s+)?(?:risen|grown|declined|fallen|dropped|climbed|surged|escalated|soared|"
    r"skyrocketed|plummeted)\b"
    r"|\b(?:has|have)\s+(?:(?:steadily|rapidly|sharply|gradually|dramatically|significantly|substantially|markedly|considerably|"
    r"greatly|almost|nearly|more\s+than)\s+)?(?:increased|decreased|doubled|tripled|quadrupled|halved)(?=\s*(?:[.,;:]|$)|\s+(?:over|"
    r"since|during|across|worldwide|globally|steadily|sharply|rapidly|dramatically|substantially|considerably|markedly|each|every|"
    r"from|in\s+(?:recent|the\s+(?:past|last))|in\s+the\s+\w+\s+(?:years|decades))\b)"
    r"|\bon\s+the\s+(?:rise|increase)\b|\bincreasingly\b"
    r"|\b(?:rising|increasing|growing|declining|decreasing|falling|upward|downward|escalating|soaring)\s+" + _TREND_NOUNS + r"\b"
    r"|\b(?:a|an)\s+(?:growing|rising|increasing)\s+(?:number|share|proportion|percentage|problem|concern|issue|burden|threat|"
    r"challenge)\b"
    # 6.6, round 58: "continues to rise", "keeps growing", "more and more", "has become more common", "ever-increasing"
    + _TREND_MORE, re.I)
_TREND_EVID = re.compile(
    r"\b(?:is|are|has\s+been|have\s+been|was|were)\s+(?:\w+\s+)?(?:rising|increasing|growing|declining|decreasing|falling|"
    r"dropping|climbing|escalating|surging|soaring)\b"
    r"|\b(?:has|have|had)\s+(?:\w+\s+)?(?:risen|grown|declined|fallen|dropped|climbed|surged|escalated|soared|increased|"
    r"decreased|doubled|tripled|halved|become|emerged|changed)\b"
    r"|\bon\s+the\s+(?:rise|increase)\b|\bincreasingly\b|\btrend(?:s|ed|ing)?\b"
    r"|\b(?:rising|increasing|growing|declining|decreasing|falling|upward|downward|escalating)\s+" + _TREND_NOUNS + r"\b"
    r"|\b(?:rose|fell|grew|declined|increased|decreased|dropped|climbed|doubled|tripled|halved)\b(?:\s+[\w%,.]+){0,6}?\s+"
    r"(?:from|between|since|over|in\s+(?:19|20)\d\d|by\s+(?:19|20)\d\d)\b"
    + _TREND_MORE +
    r"|\bover\s+(?:the\s+)?(?:past|last|previous|preceding)?\s*(?:\w+\s+){0,2}(?:years|decades|months|time)\b"
    r"|\bsince\s+(?:1[89]|20)\d\d\b|\bfrom\s+(?:1[89]|20)\d\d\s+(?:to|until|through)\b|\bbetween\s+(?:1[89]|20)\d\d\s+and\s+(?:1[89]|20)\d\d\b"
    r"|\bin\s+recent\s+(?:years|decades|times)\b|\b(?:year[- ]on[- ]year|each\s+year|every\s+year)\b"
    r"|\bbec(?:a|o)me\s+(?:more|less|increasingly|common|widespread|prevalent|popular|a\s+(?:major|growing|public|serious))\b", re.I)


def trend_ok(sentence, quoted):
    """"Is rising", "has increased", "increasingly", "a growing problem": a change over time, which a paper that measured
    one moment (a survey) does not show; only where the papers' sentences speak of a change over time."""
    return not _TREND_CLAIM.search(bare(sentence)) or bool(_TREND_EVID.search(str(quoted or "")))


# ---- how big or how wide a study was ("nationally representative", "large-scale", "multicentre")
_SCALE = {
    "representative": (r"\brepresentative\b", r"\brepresentative\b"),
    "national": (r"\bnation(?:al|wide)\s+(?:survey|study|sample|data|dataset|cohort|trial|census|database|registry|assessment|"
                 r"study|level|scale)\b|\bnation-?wide\b|\bcountry-?wide\b|\bacross\s+the\s+(?:whole\s+|entire\s+)?country\b",
                 r"\bnation(?:al|wide)\b|\bnation-?wide\b|\bcountry-?wide\b|\b(?:all|every)\s+(?:\w+\s+){0,2}(?:regions|provinces|"
                 r"states|districts|counties)\b|\bcensus\b"),
    "large": (r"\blarge[\s-]+scale\b|\b(?:a|the|this)\s+(?:very\s+)?large\s+(?:\w+\s+){0,2}?(?:study|survey|sample|cohort|trial|"
              r"dataset|population\s+study)\b|\b(?:massive|huge|extensive)\s+(?:study|survey|sample|cohort|trial|dataset)\b",
              r"\blarge\b|\b(?:[1-9]\d{0,2}(?:,\d{3})+|[1-9]\d{3,})\s+(?:\w+\s+){0,3}(?:participants|respondents|students|"
              r"pupils|patients|adults|children|adolescents|people|persons|individuals|women|men|nurses|households|subjects|users|"
              r"workers|learners|mothers|infants|cases)\b"),
    "population-based": (r"\bpopulation[\s-]+based\b", r"\bpopulation[\s-]+based\b|\bgeneral\s+population\b|\bhousehold\s+survey\b|"
                         r"\bcensus\b|\bcommunity[\s-]+based\b"),
    "multicentre": (r"\bmulti-?(?:cent(?:re|er|ric)|site|country|national|state)\b|\bmulticent\w*|\bmultisite\b",
                    r"\bmulti-?(?:cent(?:re|er|ric)|site|country|national|state)\b|\bmulticent\w*|\bmultisite\b|\b(?:\d+|two|three|"
                    r"four|five|six|seven|eight|nine|ten|several|multiple)\s+(?:\w+\s+)?(?:centres|centers|sites|hospitals|clinics|"
                    r"countries|schools|universities|regions|districts|states|provinces|facilities)\b"),
}
_SCALE_RX = {k: (re.compile(c, re.I), re.compile(e, re.I)) for k, (c, e) in _SCALE.items()}


def scale_ok(sentence, evidence):
    """How big or how wide a study was ("a nationally representative survey", "a large-scale study", "a multicentre trial")
    is what its papers say of it: one city's 412 students are not a national sample."""
    s, ev = bare(sentence), str(evidence or "")
    return all(evid.search(ev) for _k, (claim, evid) in _SCALE_RX.items() if claim.search(s))


# ---- a share in words ("one in three", "two-thirds", "a quarter of"): the papers' own share
_FRAC = {"third": 100 / 3, "quarter": 25.0, "fifth": 20.0, "tenth": 10.0}
_NUMWORD = {"one": 1, "a": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
            "twenty": 20, "hundred": 100}
_RATIO = re.compile(
    r"\b(?P<q>nearly|almost|about|around|roughly|approximately|over|more\s+than|less\s+than|under|fewer\s+than|at\s+least|"
    r"up\s+to|just\s+over|just\s+under|close\s+to|some)?\s*"
    r"(?:(?P<a>\d{1,3}|one|two|three|four|five|six|seven|eight|nine)\s+(?:in|out\s+of)\s+(?:every\s+)?"
    r"(?P<b>\d{1,4}|two|three|four|five|six|seven|eight|nine|ten|twenty|hundred)\s+(?=[a-z])"
    r"|(?P<fa>a|one|two|three|four|nine)[\s-](?P<f>thirds?|quarters?|fifths?|tenths?)\s+of\b)", re.I)
_COUNT_OF = re.compile(r"\b(\d{1,6}(?:,\d{3})?)\s+(?:of|out\s+of)\s+(?:the\s+)?(\d{1,6}(?:,\d{3})?)\b", re.I)
_OF_TOTAL = re.compile(r"\bof\s+(?:the\s+|all\s+)?(\d{1,6}(?:,\d{3})?)\s+(?:[a-z-]+\s+){0,3}?[a-z-]+,\s*(\d{1,6}(?:,\d{3})?)\b(?!\s?(?:%|percent\b|per\s+cent\b))", re.I)


def _ratio_value(m):
    if m.group("f"):
        return _NUMWORD.get(m.group("fa").lower(), 1) * _FRAC[m.group("f").lower().rstrip("s")]
    a = m.group("a").lower()
    b = m.group("b").lower()
    a = int(a) if a.isdigit() else _NUMWORD.get(a)
    b = int(b) if b.isdigit() else _NUMWORD.get(b)
    return 100.0 * a / b if a and b and a <= b else None


def _share_fits(q, p, v):
    q = re.sub(r"\s+", " ", (q or "").lower())
    if q in ("nearly", "almost", "just under", "close to"):
        return p - 6 <= v <= p + 0.5
    if q in ("over", "more than", "just over", "at least"):
        return v >= p - 0.5
    if q in ("less than", "under", "fewer than", "up to"):
        return v <= p + 0.5
    return abs(v - p) <= 3.5                 # "one in three", "about a third": within a few points of it


def ratio_ok(sentence, quoted):
    """"One in three students", "two-thirds of", "a quarter of", "nearly one in five": the share is the papers' own (their
    words, or a percentage, or "156 of 412", that is that share)."""
    q = str(quoted or "")
    pcts = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", q)]
    for a, b in _COUNT_OF.findall(q):
        a, b = int(a.replace(",", "")), int(b.replace(",", ""))
        if 0 < a <= b:
            pcts.append(100.0 * a / b)
    for b, a in _OF_TOTAL.findall(q):                  # "Of 412 students, 137 slept poorly" (6.6, round 25)
        a, b = int(a.replace(",", "")), int(b.replace(",", ""))
        if 0 < a <= b:
            pcts.append(100.0 * a / b)
    ql = _letters(q)
    for m in _RATIO.finditer(bare(sentence)):
        p = _ratio_value(m)
        if p is None:
            continue
        said = _letters(m.group(0).strip())
        if said and said in ql:
            continue
        if not any(_share_fits(m.group("q"), p, v) for v in pcts):
            return False
    return True


# ---- more words of cause ("worsens", "affects", "improves", "has a negative impact on"), outside a hedge or an aim
_CAUSE_CLAIM2 = re.compile(
    r"\b(?:worsens|harms|damages|impairs|affects|influences|drives|triggers|determines|prevents|reduces|improves|lowers|raises|"
    r"boosts|undermines|disrupts|hinders|enhances|increases\s+the\s+(?:risk|likelihood|chances?)|decreases|diminishes|stops|"
    r"eliminates|cures)\b"
    # the same verbs after a plural subject ("nets prevent infection"), never as a purpose ("used to prevent")
    # (only words that are never nouns: "harm reduction", "lower scores", "the influence of" are not claims of a cause)
    r"|(?<!\bto\s)\b(?:worsen|impair|determine|prevent|reduce|improve|undermine|disrupt|hinder|enhance|diminish|eliminate)\s+"
    r"(?!to\b|that\b|whether\b|if\b|of\b|on\b|in\b)(?=[a-z])"
    r"|\bcontribut(?:e|es|ed|ing)\s+to\b|\bprotect(?:s|ed)?\s+against\b"
    r"|\b(?:negatively|positively|adversely|significantly|strongly|directly|greatly|seriously|severely|badly|profoundly)\s+"
    r"(?:affect(?:s|ed)?|influenc(?:e|es|ed)|impact(?:s|ed)?|impair(?:s|ed)?|harm(?:s|ed)?|disrupt(?:s|ed)?|reduc(?:e|es|ed)|"
    r"improv(?:e|es|ed))\b"
    r"|\b(?:has|have|had|having|exerts?|exerted)\s+(?:a|an)\s+(?:\w+\s+){0,2}?(?:impact|effect|influence)\s+on\b", re.I)
_HEDGE = re.compile(r"\b(?:may|might|could|can|possibly|potentially|perhaps|likely\s+to|appears?\s+to|seems?\s+to|"
                    r"suggest(?:s|ed|ing)?(?:\s+that)?|whether|if|(?:examin|investigat|explor|assess|studi|understand|ask|analys|"
                    r"test|describ)\w*\s+how|to\s+(?:determine|assess|examine|investigate|explore|understand|establish|evaluate|"
                    r"identify|test|see|describe)|aim(?:s|ed)?\s+to|sought\s+to|hypothesi[sz]\w*)\b", re.I)


def _unhedged(rx, text):
    """The matches of rx in text that no hedge, question or aim ("may", "suggests that", "whether", "to examine") stands
    before (within the eight words before it, in the same clause)."""
    out = []
    for m in rx.finditer(text):
        before = re.split(r"[,;:]|\bbut\b|\bwhereas\b|\bwhile\b|\balthough\b", text[max(0, m.start() - 90):m.start()])[-1]
        if not _HEDGE.search(" ".join(before.split()[-8:])):
            out.append(m)
    return out


def cause2_ok(sentence, quoted):
    """"Social media use worsens sleep", "has a negative impact on" stand only on papers whose sentences speak of a cause,
    an effect or a trial, as cause_ok; "may worsen", "whether it affects", "to examine how it affects" are not claims."""
    if not _unhedged(_CAUSE_CLAIM2, bare(sentence)):
        return True
    q = str(quoted or "")
    return bool(_CAUSE_EVID.search(q)) or not _LINK_EVID.search(q)


# ---- a hedge kept: the papers say "may contribute to", the sentence may not say "contributes to"
_HEDGED_VERBS = re.compile(r"\b(contribut|caus|lead|led|result|increas|decreas|reduc|improv|wors|affect|influenc|predict|protect|"
                           r"prevent|explain|driv|trigger|harm|impair|benefit|enhanc|lower|rais)\w*", re.I)


def hedge_ok(sentence, quoted):
    """When a paper says a thing only with a hedge ("may contribute to", "could reduce", "appears to improve", "suggests
    that it affects"), a sentence that says it without one ("contributes to") says more than the paper does."""
    plain = {m.group(1).lower()[:5] for m in _unhedged(_HEDGED_VERBS, bare(sentence))}
    if not plain:
        return True
    q = str(quoted or "")
    q_plain = {m.start() for m in _unhedged(_HEDGED_VERBS, q)}
    for stem in plain:
        said = [m for m in _HEDGED_VERBS.finditer(q) if m.group(1).lower()[:5] == stem]
        if said and not any(m.start() in q_plain for m in said):
            return False
    return True


# ---- "a recent study" for an old paper; "the first study to"
_RECENT = re.compile(r"\brecent(?:ly)?\s+(?:published\s+)?(?:study|studies|research|review|reviews|survey|surveys|trial|trials|paper|"
                     r"papers|work|evidence|findings|data|report|reports|analysis|analyses|investigation|investigations|literature)\b"
                     r"|\b(?:was|were|has\s+been|have\s+been)\s+recently\s+(?:shown|found|reported|published|demonstrated|documented)\b"
                     r"|\brecently\s+(?:shown|found|reported|published|demonstrated|documented)\b", re.I)
RECENT_YEARS = 5           # "recent" = published in the last five years (the span most reviews and supervisors ask for)


def recent_ok(sentence, pub_years=(), this_year=None):
    """"A recent study", "recently reported": every paper it cites is from the last five years."""
    if not _RECENT.search(bare(sentence)):
        return True
    ys = [int(y) for y in (pub_years or ()) if str(y or "").isdigit()]
    if not ys:
        return True
    import datetime as _dt
    ty = int(this_year or _dt.datetime.now(_dt.timezone.utc).year)
    return min(ys) >= ty - RECENT_YEARS


_FIRST = re.compile(r"\b(?:the\s+)?first\s+(?:study|studies|trial|survey|review|paper|research|investigation|analysis|report|"
                    r"evidence)\b|\bfor\s+the\s+first\s+time\b|\bthe\s+first\s+to\b", re.I)


def first_ok(sentence, quoted):
    """"The first study to…", "for the first time": only where the papers' sentences say "first"."""
    return not _FIRST.search(bare(sentence)) or bool(re.search(r"\bfirst\b", str(quoted or ""), re.I))


# ---- 6.6, round 21: what research has NOT done ("no study has…", "little is known", "few studies", "a paucity of research")
# is a claim about all the literature: one paper's abstract can back it only when it says so itself
_ABSENCE = re.compile(r"\bno\s+(?:previous\s+|prior\s+|other\s+|published\s+)?(?:stud(?:y|ies)|research|data|evidence|investigations?)\s+(?:has|have|had|exists?|is|are)\b|"
                      r"\bha(?:s|ve)\s+(?:not|never)\s+(?:yet\s+)?been\s+(?:studied|examined|investigated|explored|researched|assessed|reported|documented|described)\b|"
                      r"\blittle\s+(?:is|was)\s+known\b|\blittle\s+(?:research|evidence|attention)\b|\bfew\s+(?:stud(?:y|ies)|researchers?)\s+(?:ha(?:s|ve)|exist|examined|"
                      r"investigated|explored|looked|focused)\b|\bpaucity\s+of\b|\bdearth\s+of\b|\b(?:scarce|scant|sparse|limited)\s+(?:research|evidence|data|"
                      r"literature|studies)\b|\b(?:research|evidence|data|literature|studies)\b[^.;]{0,60}?\b(?:is|are|remains?)\s+(?:scarce|scant|sparse|limited|lacking)\b|"
                      r"\bgap\s+in\s+(?:the\s+)?(?:literature|knowledge|research|evidence)\b|\bunder-?(?:studied|researched|explored)\b", re.I)
# (6.6, round 26) "no study has", "has not been studied" say there is NONE: only a paper that says none ("no study",
# "has not been examined", "the first study") backs them; "little is known", "few studies" are backed by any gap the paper
# states. A paper's "no association" says nothing of what research hasn't done.
_ABSENCE_ALL = re.compile(r"\bno\s+(?:previous\s+|prior\s+|other\s+|published\s+)?(?:stud(?:y|ies)|research|data|evidence|investigations?)\s+"
                          r"(?:has|have|had|exists?|is|are)\b|\bha(?:s|ve)\s+(?:not|never)\s+(?:yet\s+)?been\s+(?:studied|examined|investigated|"
                          r"explored|researched|assessed|reported|documented|described)\b", re.I)
_ABSENCE_EVID_ALL = re.compile(r"\bno\s+(?:previous\s+|prior\s+|other\s+|published\s+|earlier\s+)?(?:stud(?:y|ies)|research|data|evidence|investigations?|"
                               r"surveys?|trials?)\b|\b(?:not|never)\s+(?:yet\s+|previously\s+|before\s+)?been\s+(?:studied|examined|investigated|explored|"
                               r"researched|assessed|reported|documented|described)\b|\bthe\s+first\s+(?:\w+\s+){0,2}?(?:study|survey|trial|review|"
                               r"investigation|analysis|report)\b|\bfor\s+the\s+first\s+time\b", re.I)
_ABSENCE_EVID = re.compile(_ABSENCE_EVID_ALL.pattern + r"|\bfew\s+(?:stud(?:y|ies)|researchers?|data|trials?|surveys?)\b|\blittle\s+(?:is|was)\s+known\b|"
                           r"\blittle\s+(?:research|evidence|attention|data)\b|\b(?:limited|scarce|scant|sparse)\s+(?:research|evidence|data|literature|"
                           r"studies)\b|\b(?:research|evidence|data|literature|studies)\b[^.;]{0,60}?\b(?:is|are|remains?)\s+(?:scarce|scant|sparse|"
                           r"limited|lacking)\b|\black(?:ing)?\s+(?:of\s+)?(?:research|studies|data|evidence)\b|\bpaucity\b|\bdearth\b|"
                           r"\bgaps?\s+in\s+(?:the\s+)?(?:literature|knowledge|research|evidence)\b|\bunder-?(?:studied|researched|explored)\b|"
                           r"\b(?:unknown|unclear|poorly\s+understood|not\s+well\s+(?:understood|known|studied))\b", re.I)


def absence_ok(sentence, quoted):
    """"No study has examined…", "little is known about…", "few studies…", "a gap in the literature": a claim about what
    research has not done. Only where the papers' own sentences say the same kind of thing (a paper's statement of the
    gap it fills); a paper that simply studied something says nothing about what others haven't. "No study has" and "has
    not been studied" say there is none: the paper must say that itself ("little is known" is not "no study"), 6.6 round 26."""
    s, q = bare(sentence), str(quoted or "")
    if _ABSENCE_ALL.search(s):
        return bool(_ABSENCE_EVID_ALL.search(q))
    return not _ABSENCE.search(s) or bool(_ABSENCE_EVID.search(q))


# ---- 6.6, round 22: a correlation's strength in words keeps to its size ("a strong correlation" is not r = .23)
_STRENGTH_SAID = re.compile(r"\b(strong(?:ly)?|weak(?:ly)?|moderate(?:ly)?|modest(?:ly)?|small|large|high|low)\s+(?:(?:and\s+)?(?:positive(?:ly)?|negative(?:ly)?|"
                            r"inverse(?:ly)?|significant(?:ly)?|statistically\s+significant)\s+)?(?:correlat\w*|associat\w*|relationship|relation|link(?:ed)?)\b", re.I)
_R_VALUE = re.compile(r"(?<![A-Za-z])(?:r|ρ|rho|r_?s|rs|Pearson['’]?s\s+r|Spearman['’]?s\s+(?:rho|ρ))\s*(?:\(\s*\d+\s*\))?\s*=\s*([-−–]?\s?0?\.\d+)", re.I)


def _band(word):
    w = word.lower()
    return "strong" if w.startswith(("strong", "large", "high")) else "weak" if w.startswith(("weak", "small", "low", "modest")) else "moderate"


def strength_ok(sentence, quoted):
    """"A strong correlation" (or weak, moderate) keeps to the size the papers give, by the usual reading of r (weak below
    .3, moderate .3 to .5, strong .5 and above), with room either side; unless the paper calls it that itself."""
    s = bare(sentence)
    m = _STRENGTH_SAID.search(s)
    if not m:
        return True
    q = str(quoted or "")
    word = m.group(1).lower()
    if re.search(r"\b" + re.escape(word[:6]), q, re.I):
        return True                                  # (the paper's own word for it)
    said = {_band(x.group(1)) for x in _STRENGTH_SAID.finditer(q)}
    if said and _band(word) not in said:
        return False                                 # (6.6, round 23: the paper calls it moderate; the sentence, strong)
    vals = []
    for v in _R_VALUE.findall(q):
        try:
            vals.append(abs(float(re.sub(r"\s+", "", v).replace("−", "-").replace("–", "-"))))
        except ValueError:
            pass
    if not vals:
        return True                                  # (no r to compare with: other rules hold the sentence to its paper)
    band = _band(word)
    ok = {"strong": lambda r: r >= 0.4, "weak": lambda r: r < 0.4, "moderate": lambda r: 0.2 <= r < 0.7}[band]
    return any(ok(r) for r in vals)


# ---- a finding's polarity ("associated with" where the paper found no association, and the other way round)
# ---- 6.6, round 101 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a level, or a share, said the other way. A survey of knowledge, attitudes and practice (and of uptake,
# adherence, satisfaction, quality…) reports levels in words: "knowledge was poor", "positive attitudes", "low uptake". "Knowledge
# was high" where the paper says it was low, or "good knowledge" for its "poor knowledge", states the opposite of the finding;
# so does "38% of students did not sleep enough" where 38% slept enough (or "were vaccinated" for "were unvaccinated").
_LEVEL_POS = (r"high|good|adequate|sufficient|satisfactory|favou?rable|positive|excellent|appropriate|correct|acceptable|optimal|"
              r"above[\s-]average")
_LEVEL_NEG = (r"low|poor|inadequate|insufficient|unsatisfactory|unfavou?rable|negative|bad|inappropriate|incorrect|unacceptable|"
              r"suboptimal|sub-optimal|below[\s-]average|limited|deficient")
_LEVEL_TOPIC = (r"knowledge|awareness|attitudes?|practices?|perceptions?|adherence|compliance|uptake|utili[sz]ation|usage|coverage|"
                r"satisfaction|quality|performance|self-?efficacy|confidence|motivation|engagement|participation|acceptability|"
                r"acceptance|willingness|readiness|literacy|competence|skills?|understanding|beliefs?|access|support")
_LV_POLE = r"(?P<pole>" + _LEVEL_POS + "|" + _LEVEL_NEG + r")"
_LV_TOPIC = r"(?P<topic>" + _LEVEL_TOPIC + r")"
_LV_NOT = re.compile(r"\b(?:not|never|no|n['’]t|nor|neither)\b|n['’]t\b", re.I)
_LV_RX = [re.compile(r"\b" + _LV_POLE + r"\s+(?:[a-z][\w-]*\s+){0,2}?" + _LV_TOPIC + r"\b", re.I),
          re.compile(r"\b" + _LV_POLE + r"\s+levels?\s+of\s+(?:[a-z][\w-]*\s+){0,2}?" + _LV_TOPIC + r"\b", re.I),
          re.compile(r"\b" + _LV_TOPIC + r"\b(?P<mid>(?:\s+[\w’'-]+){0,8}?)\s+(?:was|were|is|are|remained|remains|appeared|seemed|"
                     r"(?:was|were)\s+found\s+to\s+be|(?:was|were)\s+(?:considered|rated|described\s+as|judged))\s+"
                     r"(?P<adv>(?:\w+ly|very|quite|fairly|rather|largely|mostly)\s+){0,2}" + _LV_POLE + r"\b(?!\s*[-–]?\s*(?:to|and|or)\s+(?:moderate|"
                     + _LEVEL_POS + "|" + _LEVEL_NEG + r")\b)", re.I)]


def _lv_pole(w):
    return "pos" if re.fullmatch(_LEVEL_POS, w.lower()) else "neg"


def _lv_topic(w):
    w = w.lower().replace("utilis", "utiliz").replace("-", "")
    return w[:6]


def _levels(text):
    """{(topic, "pos"|"neg")}: "poor knowledge", "high levels of awareness", "adherence to treatment was low"."""
    out = set()
    t = re.sub(r"\s+", " ", str(text or ""))
    for rx in _LV_RX:
        for m in rx.finditer(t):
            span = t[m.start():m.end()]
            if re.search(r"[;:]|\bbut\b|\bwhereas\b|\bwhile\b|\bthan\b", span, re.I):
                continue                    # (another clause in between: not this topic's level)
            before = t[max(0, m.start() - 14):m.start()]
            pole_at = m.start("pole")
            if t[max(0, pole_at - 1):pole_at] == "-" or re.match(r"(?:\s+|-)(?:school|income|risk|birth|blood|grade|resource|middle|"
                                                                  r"lying|cost|dose)\b", t[m.end("pole"):], re.I):
                continue                    # ("HIV-positive", "high school", "low-income", "low birth weight": not a level)
            if _LV_NOT.search(span) or re.search(r"\b(?:not|no|n['’]t)\s*$", before, re.I):
                continue                    # ("not good", "no poor"): left to the checker
            if re.search(r"\b(?:" + _LEVEL_POS + "|" + _LEVEL_NEG + r")\b", (m.groupdict().get("mid") or ""), re.I):
                continue                    # ("knowledge of good practice was poor": the level is the last word's)
            out.add((_lv_topic(m.group("topic")), _lv_pole(m.group("pole"))))
    return out


def level_ok(sentence, quoted):
    """A level said as the papers say it: when the sentence calls something high (good, adequate, positive…) and the papers'
    sentences call that same thing low (poor, inadequate, negative…), and never as the sentence does, it is turned over."""
    s = _levels(bare(sentence))
    if not s:
        return True
    q = _levels(quoted)
    for topic, pole in s:
        if (topic, "neg" if pole == "pos" else "pos") in q and (topic, pole) not in q:
            return False
    return True


_SG_NEG = re.compile(r"^(?:not|never|no|none|nor|neither|without|unable|lacked|lacks|lack)$|n['’]t$", re.I)
_SG_SKIP = {"of", "the", "a", "an", "and", "or", "to", "in", "on", "at", "for", "with", "by", "was", "were", "is", "are", "had", "has", "have",
            "did", "do", "does", "been", "be", "who", "which", "that", "they", "them", "their", "those", "these", "this", "it", "its", "all",
            "respondents", "participants", "students", "pupils", "people", "patients", "nurses", "women", "men", "children", "adults",
            "households", "mothers", "workers", "only", "about", "nearly", "almost", "over", "some", "also", "reported", "said", "indicated",
            "stated", "agreed", "would", "could", "can", "will", "there", "than", "more", "less"}


def _sg_clauses(text):
    """{figure: [(negated, {stems}, {un-words})]}: each percentage and what is said of it, up to the clause's end."""
    out = {}
    t = re.sub(r"\bnot\s+(?:only|just|merely|less|fewer)\b|\bno\s+(?:longer|less|more|fewer)\b", " ", str(text or ""), flags=re.I)
    for m in re.finditer(r"(?<![\d.])(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", t):
        after = re.split(r"[,;:()]|\.(?!\d)|\b(?:but|while|whereas|compared|than|respectively)\b|\band\s+\d", t[m.end():m.end() + 140], flags=re.I)[0]
        # (6.6, round 111: and the words before it in its clause, for a share said in the passive: "a net was not used by 38%")
        before = re.split(r"[,;:()]|\.(?!\d)|\b(?:but|while|whereas|compared|than|respectively|and)\b|\d", t[max(0, m.start() - 120):m.start()], flags=re.I)[-1]
        if not re.search(r"\bby\s*$", before, re.I):
            before = ""
        neg, stems, un = False, set(), set()
        for w in re.findall(r"[A-Za-z][A-Za-z'’-]*", before)[-10:] + re.findall(r"[A-Za-z][A-Za-z'’-]*", after)[:12]:
            lw = w.lower()
            if _SG_NEG.search(lw):
                neg = not neg
                continue
            if lw in _SG_SKIP or len(lw) < 3:
                continue
            stems.add(lw[:4])
            if lw.startswith("un") and len(lw) >= 7:
                un.add(lw[2:6])
        out.setdefault(float(m.group(1)), []).append((neg, stems, un))
    return out


def share_sign_ok(sentence, quoted):
    """A share keeps its "not": "38% of students did not sleep enough" where the papers' sentences give 38% as those who
    slept enough (and never as those who didn't), or "were vaccinated" for their "were unvaccinated", is turned over."""
    s = _sg_clauses(bare(sentence))
    if not s:
        return True
    q = _sg_clauses(quoted)
    for v, mine in s.items():
        theirs = q.get(v) or []
        for neg, st, un in mine:
            if not st:
                continue
            same, other = False, False
            for neg2, st2, un2 in theirs:
                a_neg, b_neg = neg, neg2
                a, b = set(st), set(st2)
                for u in un & b:              # "unvaccinated" here, "vaccinated" there
                    a_neg, a = not a_neg, (a - {"un" + u[:2]}) | {u}
                for u in un2 & a:
                    b_neg, b = not b_neg, (b - {"un" + u[:2]}) | {u}
                shared = a & b
                if not shared or len(shared) / len(a | b) < 0.5:
                    continue
                if a_neg == b_neg:
                    same = True
                else:
                    other = True
            if other and not same:
                return False
    return True


# ---- 6.6, round 102 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): which side of a comparison is the higher, and
# which way a link goes. "Girls slept more than boys" where the papers' sentences say boys slept more than girls, "the
# intervention group slept longer than the control group" for the opposite, "positively associated" for "negatively", "a risk
# factor" for "protective": each states the other side of the finding, and none was caught. Sides are read from pairs people
# compare (girls and boys, urban and rural, intervention and control…); a comparison said in other words is left to the checker.
_CMP_PAIRS = [(r"girls?", r"boys?"), (r"wom[ae]n", r"m[ae]n"), (r"females?", r"males?"), (r"urban", r"rural"), (r"older", r"younger"),
              (r"(?:intervention|experimental|treatment)(?:\s+(?:group|arm))?", r"control(?:\s+(?:group|arm))?|comparison\s+group"),
              (r"first[-\s]year", r"(?:final|fourth|third|last)[-\s]year"), (r"public", r"private"), (r"married", r"unmarried|single"),
              (r"employed", r"unemployed"), (r"smokers", r"non-?smokers"), (r"users", r"non-?users"), (r"cases", r"controls"),
              (r"exposed", r"unexposed|non-?exposed"), (r"literate", r"illiterate"), (r"(?:weekdays|school\s+(?:nights|days))", r"weekends?"),
              (r"day\s+students", r"boarders|boarding\s+students"), (r"insured", r"uninsured"), (r"vaccinated", r"unvaccinated"),
              (r"(?:the\s+)?(?:pre-?test|baseline)", r"(?:the\s+)?(?:post-?test|follow-?up|endline)"),
              (r"after", r"before")]                       # (6.6, round 102: "higher after the training than before")
_CMP_UP = r"more\s+likely|more|higher|greater|larger|longer|better|bigger|stronger|faster"
_CMP_DOWN = r"less\s+likely|less|fewer|lower|smaller|shorter|worse|weaker|slower"
_CMP_WORD = re.compile(r"\b(?P<c>" + _CMP_UP + "|" + _CMP_DOWN + r")\b", re.I)
_CMP_STOP = {"likely", "to", "than", "of", "the", "a", "an", "in", "among", "amongst", "for", "on", "at", "and", "was", "were", "is", "are",
             "had", "have", "has", "did", "do", "does", "be", "been", "their", "they", "who", "that", "significantly", "statistically",
             "slightly", "much", "far", "considerably", "substantially", "markedly", "somewhat", "times", "also", "both", "it", "its"}


def _cmp_side(text, pair):
    """"a", "b" or None: which of the pair the text names (one only)."""
    a = bool(re.search(r"(?<![\w-])(?:" + pair[0] + r")(?![\w-])", text, re.I))
    b = bool(re.search(r"(?<![\w-])(?:" + pair[1] + r")(?![\w-])", text, re.I))
    return "a" if a and not b else "b" if b and not a else None


def _cmp_words(text, pairs_rx):
    """The content words of text, lower-cased, without the pairs' own words or names (a capital inside the text: a place)."""
    out = []
    for k, m in enumerate(re.finditer(r"[A-Za-z][A-Za-z-]*", text)):
        w = m.group(0)
        if w.lower() in _CMP_STOP or pairs_rx.fullmatch(w) or (w[0].isupper() and m.start() > 0 and text[:m.start()].strip()):
            continue
        out.append(w.lower())
    return out


_CMP_ALL = re.compile("|".join("(?:" + a + ")|(?:" + b + ")" for a, b in _CMP_PAIRS) + r"|students?|pupils?|nurses?|patients?|participants?|"
                      r"respondents?|groups?|arms?|people|children|adults|adolescents|those|ones|counterparts", re.I)


def _comparisons(text):
    """{(pair index, side that is higher, measure stems)}: "girls slept more than boys", "prevalence was higher among
    urban than rural students", "more girls than boys used phones", "scores were lower at follow-up than at baseline"."""
    out = set()
    t = re.sub(r"\s+", " ", str(text or ""))
    for m in _CMP_WORD.finditer(t):
        rest = t[m.end():m.end() + 90]
        than = re.search(r"\bthan\b", rest)
        if not than or re.search(r"[;:]|\.(?!\d)", rest[:than.start()]):
            continue
        clause = re.split(r"[;:]|,\s*(?:but|while|whereas)\b|\b(?:but|whereas|while|although)\b|\.(?!\d)|,\s*and\b", t[:m.start()])[-1][-90:]
        mid = rest[:than.start()]
        tail = re.split(r"[,;:]|\.(?!\d)|\b(?:but|while|whereas)\b|\band\s+(?!\w+\s+than\b)", rest[than.end():])[0][:80]
        up = bool(re.fullmatch(_CMP_UP, m.group("c"), re.I))
        to = re.search(r"\bto\s+([A-Za-z][A-Za-z-]*)", tail)
        measure = {w[:4] for w in _cmp_words(mid, _CMP_ALL)[:2] + _cmp_words(clause, _CMP_ALL)[-1:]
                   + ([to.group(1).lower()] if to else []) + _cmp_words(tail, _CMP_ALL)[:3]}
        if not measure:
            continue
        for i, pair in enumerate(_CMP_PAIRS):
            o = _cmp_side(tail, pair)
            if not o:
                continue
            bare_mid = not _cmp_words(mid, _CMP_ALL) and len(mid.split()) <= 3
            lead = re.match(r"\s*(?:the\s+)?(?:" + pair[0] + "|" + pair[1] + r")\b", mid, re.I)          # ("higher after the training than before")
            sub = _cmp_side(mid, pair) if (bare_mid or lead or re.search(r"\b(?:among|amongst|in|for|at|on|during)\b", mid, re.I)) else None
            sub = sub or _cmp_side(clause, pair)
            if not sub or sub == o:
                continue
            out.add((i, sub if up else o, frozenset(measure)))
    return out


def comparison_ok(sentence, quoted):
    """A comparison with the papers' higher side: when the sentence says one of a pair (girls, urban, the intervention group…)
    is higher, more or better than the other, and the papers' sentences compare the same pair on the same thing the other
    way, and never this way, it is turned over."""
    s = _comparisons(bare(sentence))
    if not s:
        return True
    q = _comparisons(quoted)
    for i, hi, ms in s:
        same = [h for j, h, m2 in q if j == i and len(ms & m2) * 2 >= len(ms | m2)]      # (the same thing compared)
        if same and hi not in same:
            return False
    return True


_SIGN = re.compile(r"\b(?P<w>positive(?:ly)?|direct(?:ly)?|negative(?:ly)?|inverse(?:ly)?)\s+(?:and\s+)?(?:(?:statistically\s+)?significant(?:ly)?\s+)?"
                   r"(?:associat\w*|correlat\w*|relat\w*|link\w*)|"
                   r"\b(?:associat\w*|correlat\w*|related|linked)\s+(?:(?:statistically\s+)?significantly\s+)?(?P<w2>positively|negatively|inversely|directly)\b", re.I)


def _signs(text):
    out = set()
    for m in _SIGN.finditer(str(text or "")):
        w = (m.group("w") or m.group("w2") or "").lower()
        out.add("+" if w.startswith(("pos", "dir")) else "-")
    return out


def sign_ok(sentence, quoted):
    """A link's sign as the papers give it: "positively associated" where they say negatively (or inversely), and never
    positively, is turned over; and "a risk factor" where they say protective (or the other way)."""
    s, q = _signs(bare(sentence)), _signs(quoted)
    if s and q and not (s & q):
        return False
    b = bare(sentence)
    neg = re.compile(r"\b(?:lack|absence|not|non|no|without|never|low|poor|inadequate|insufficient|short|reduced)\b", re.I)
    risk_s, prot_s = bool(re.search(r"\brisk\s+factors?\b", b, re.I)), bool(re.search(r"\bprotective\b", b, re.I))
    risk_q, prot_q = bool(re.search(r"\brisk\s+factors?\b", str(quoted or ""), re.I)), bool(re.search(r"\bprotective\b", str(quoted or ""), re.I))
    if (risk_s != prot_s) and (risk_q != prot_q) and risk_s != risk_q and not neg.search(b) and not neg.search(str(quoted or "")):
        return False
    return True


# ---- 6.6, round 103 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): whose is the superlative. "Malaria was the most
# common cause of admission" where the papers' sentences give that place to pneumonia ("pneumonia was the most common cause,
# followed by malaria"), or "smoking was the strongest predictor" for their caffeine, gives the finding to the wrong thing.
_SUP_WORDS = (r"most\s+(?:common|frequent|prevalent|important|significant|cited|reported|used|preferred|popular|influential)|commonest|"
              r"leading|main|major|primary|principal|strongest|biggest|largest|greatest|highest|chief|predominant|dominant")
_SUP_ANY = re.compile(r"\b(?:" + _SUP_WORDS + r")\b", re.I)
_SUP_STOP = {"the", "a", "an", "of", "in", "on", "at", "for", "and", "or", "among", "amongst", "with", "by", "to", "this", "that", "these",
             "those", "their", "its", "his", "her", "overall", "also", "still", "both", "all", "only", "however", "thus", "therefore"}
_SUP_A = re.compile(r"(?P<subj>[^,;:()]{1,80}?)\s+(?:was|were|is|are|remains|remained|became|emerged\s+as|ranked\s+as)\s+(?:by\s+far\s+)?the\s+"
                    r"(?:single\s+)?(?P<sup>" + _SUP_WORDS + r")(?:\s+(?P<noun>[a-z][\w-]*))?", re.I)
_SUP_B = re.compile(r"\bthe\s+(?:single\s+)?(?P<sup>" + _SUP_WORDS + r")\s+(?P<noun>[a-z][\w-]*)(?:\s+of\s+[\w\s-]{1,40}?)?\s+(?:was|were|is|are)\s+"
                    r"(?P<subj>[^,;:()]{1,60}?)(?=\s*(?:[,;:().]|$|\s+followed\b|\s+and\s+then\b|\s+\())", re.I)


def _sup_subj(words):
    return {w.lower()[:4] for w in re.findall(r"[A-Za-z][\w-]*", words) if w.lower() not in _SUP_STOP and len(w) >= 3}


def _superlatives(text):
    """{(the superlative and its noun, the stems of what it is said of)}: "malaria was the most common cause", "the leading
    cause of death was pneumonia"."""
    out = set()
    t = re.sub(r"\s+", " ", str(text or ""))
    for clause in re.split(r"[;]|\.(?!\d)\s", t):
        if not _SUP_ANY.search(clause):
            continue                            # (6.6, round 115: only a clause with a superlative is read)
        for rx in (_SUP_A, _SUP_B):
            for m in rx.finditer(clause):
                subj = m.group("subj")
                if rx is _SUP_A:
                    subj = re.split(r",|\bthat\b|\bwhich\b|\bwho\b", subj)[-1]
                    subj = " ".join(subj.split()[-4:])
                stems = _sup_subj(subj)
                if not stems or re.search(r"\b(?:not|never|no)\b", m.group(0), re.I):
                    continue
                key = re.sub(r"\s+", " ", m.group("sup").lower()) + "|" + ((m.group("noun") or "").lower()[:4])
                out.add((key, frozenset(stems)))
    return out


def superlative_whose_ok(sentence, quoted):
    """A superlative given to what the papers give it to: when the sentence says X is the most common (the leading, the
    strongest…) of something and the papers' sentences say that of something else, and never of X, it is not theirs."""
    s = _superlatives(bare(sentence))
    if not s:
        return True
    q = _superlatives(quoted)
    for key, subj in s:
        theirs = [sj for k, sj in q if k == key]
        if theirs and not any(subj & sj for sj in theirs):
            return False
    return True


# ---- 6.6, round 106 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a figure from one place is that place's.
# Research writing goes from the world down to the country and the place (the owner's 15 Nov 12:46), so a country's figure is
# asked for; "In Ghana, 38% of nursing students self-medicate (Mensah, 2021)" where Mensah studied one college in Kumasi says
# of a whole country what was found in one college. A figure said of a country (with no narrower place in the sentence) is
# kept only where its paper is national too (a national or nationally representative survey, a DHS, every region…) or names
# no narrower setting.
_COUNTRIES = set(_PLACE_NAMES[:_PLACE_NAMES.index("africa")]) - {"hong kong", "puerto rico"}
_NATIONAL = re.compile(r"\bnational(?:ly)?\b|\bnation-?wide\b|\bcountry-?wide\b|\bacross\s+(?:the\s+)?(?:country|nation)\b|"
                       r"\bdemographic\s+and\s+health\s+surveys?\b|\bDHS\b|\bMICS\b|\bcensus\b|\b(?:all|every|each)\s+(?:the\s+)?(?:\d+\s+|"
                       r"(?:ten|sixteen|twelve|thirty-six|\w+)\s+)?(?:regions|states|provinces|counties)\b|\brepresentative\s+sample\b", re.I)
_NARROW = re.compile(r"\b(?:a|an|one|two|three|four|five|\d+|the|selected|some|several)\s+(?:(?:large|small|single|rural|urban|peri-?urban|public|"
                     r"private|government|mission|faith-?based|teaching|tertiary|secondary|primary|district|regional|municipal|"
                     r"metropolitan|senior\s+high|junior\s+high|basic|nursing|midwifery|health|selected)\s+){0,3}"
                     r"(?:hospitals?|universit(?:y|ies)|colleges?|schools?|clinics?|polyclinics?|health\s+(?:centres?|centers?|facilit(?:y|ies))|"
                     r"facilit(?:y|ies)|communit(?:y|ies)|districts?|municipalit(?:y|ies)|metropolis|villages?|towns?|wards?|campus(?:es)?|"
                     r"sub-?districts?|prisons?|markets?|churches|mosques|households\s+in)\b|"
                     r"\b[A-Z][a-z]+(?:[\s-][A-Z][a-z]+){0,3}\s+(?:Teaching\s+Hospital|Hospital|University|College|Municipality|Municipal|"
                     r"District|Metropolis|Metropolitan|Region|Polyclinic|Clinic|Senior\s+High\s+School|School)\b")


def _countries_in(text):
    out = set()
    for m in _PLACE_RX.finditer(str(text or "")):
        p = m.group(1).lower()
        if p in _COUNTRIES and not (p in _CAPS_ONLY and not m.group(1)[0].isupper()):
            out.add(p)
    for m in _DEMONYM_RX.finditer(str(text or "")):
        p = _DEMONYMS.get(m.group(1).lower())
        if p in _COUNTRIES:
            out.add(p)
    return out


def _local_in(text):
    """A city or a narrower setting named in the text."""
    t = str(text or "")
    if _NARROW.search(t):
        return True
    return any(m.group(1).lower() not in _COUNTRIES and m.group(1).lower() not in _REGIONS_WIDE for m in _PLACE_RX.finditer(t))


_REGIONS_WIDE = set(_PLACE_NAMES[_PLACE_NAMES.index("africa"):_PLACE_NAMES.index("accra")])


def scope_ok(sentence, evidence):
    """A figure said of a whole country stands only where its paper is national too, or names no narrower setting."""
    s = bare(sentence)
    if not re.search(r"\d", s) and not re.search(r"\b(?:half|third|quarter|majority|most|one\s+in)\b", s, re.I):
        return True
    if not _countries_in(s) or _local_in(s) or _NATIONAL.search(s) is None and re.search(r"\b(?:a|one)\s+(?:study|survey)\b", s, re.I):
        return True                  # (no country, a narrower place named, or "a study in Ghana": the study's place, not the country's figure)
    ev = str(evidence or "")
    if _NATIONAL.search(ev):
        return True
    return not _local_in(ev)


# ---- 6.6, round 108 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): how often keeps its unit. "38% of students used
# their phones daily after midnight" where the paper's 38% used them weekly (or "drank every day" for "at least once a month")
# says something else of the same figure. Compared for the same percentage, in the words that follow it.
_FREQ = (("day", re.compile(r"\bdaily\b|\bnightly\b|\bevery\s+(?:day|night|morning|evening)\b|\beach\s+(?:day|night)\b|"
                            r"\b(?:once|twice|thrice|\w+\s+times)\s+(?:a|per|each)\s+(?:day|night)\b|\bper\s+day\b|\bday\s+after\s+day\b", re.I)),
         ("week", re.compile(r"\bweekly\b|\bevery\s+week\b|\beach\s+week\b|\b(?:once|twice|thrice|\w+\s+times)\s+(?:a|per|each)\s+week\b|"
                             r"\bper\s+week\b|\bevery\s+weekend\b", re.I)),
         ("month", re.compile(r"\bmonthly\b|\bevery\s+month\b|\beach\s+month\b|\b(?:once|twice|thrice|\w+\s+times)\s+(?:a|per|each)\s+month\b|"
                              r"\bper\s+month\b", re.I)),
         ("year", re.compile(r"\byearly\b|\bannually\b|\bevery\s+year\b|\beach\s+year\b|\b(?:once|twice|thrice|\w+\s+times)\s+(?:a|per|each)\s+year\b|"
                             r"\bper\s+(?:year|annum)\b", re.I)),
         ("ever", re.compile(r"\b(?:ever|never)\b", re.I)))


def _freq_clauses(text):
    """{figure: [{units}]}: each percentage and how often the words after it say (daily, weekly, monthly, yearly, ever)."""
    out = {}
    t = str(text or "")
    for m in re.finditer(r"(?<![\d.])(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", t):
        after = re.split(r"[,;:()]|\.(?!\d)|\b(?:but|while|whereas|compared|than|respectively)\b|\band\s+\d", t[m.end():m.end() + 140], flags=re.I)[0]
        units = {name for name, rx in _FREQ if rx.search(after)}
        if units:
            out.setdefault(float(m.group(1)), []).append(units)
    return out


def frequency_ok(sentence, quoted):
    """The same share said with another "how often" than the papers' (daily for their weekly) is not theirs."""
    s = _freq_clauses(bare(sentence))
    if not s:
        return True
    q = _freq_clauses(quoted)
    for v, mine in s.items():
        theirs = q.get(v) or []
        if not theirs:
            continue
        for u in mine:
            if not any(u & t for t in theirs):
                return False
    return True


# ---- 6.6, round 109 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a share keeps the people it narrows to. "38% of pregnant women were anaemic" where the paper's 38% is of
# all the women it studied (or "38% of women" for its "38% of pregnant women") gives the figure to other people; round 19's
# check took any shared word ("women") as the same people.
_RESTRICT = re.compile(r"\b(pregnan\w*|lactating|breast-?feeding|nursing\s+mothers|(?:un)?married|single|divorced|widowed|"
                       r"hiv[\s-]*(?:positive|negative|infected)|living\s+with\s+hiv|diabetic|hypertensive|disabled|elderly|older|younger|"
                       r"adolescents?|teenage|sexually\s+active|first-?time|primiparous|multiparous|post-?natal|post-?partum|antenatal|"
                       r"in-?school|out-?of-?school|school-?going|refugees?|migrants?|homeless|street)\b", re.I)


_RESTRICT_SAME = {"antenata": "pregnan", "pregnanc": "pregnan", "pregnant": "pregnan", "breastfe": "lactatin", "nursing ": "lactatin",
                  "living w": "hiv posi", "hiv infe": "hiv posi", "teenage": "adolesce"}


def _restricting(words):
    out = set()
    for m in _RESTRICT.finditer(" ".join(words or [])):
        k = re.sub(r"[\s-]+", " ", m.group(1).lower()).rstrip("s")[:8]
        out.add(_RESTRICT_SAME.get(k, k))     # (women attending antenatal care are pregnant women)
    return out


def narrowing_ok(sentence, quoted):
    """The same "N% of X" keeps X's narrowing words: the papers' "pregnant women" (married, HIV-positive, adolescent…) is not
    "women", nor are their women pregnant ones unless the papers say so."""
    said, ev = _shares_of(bare(sentence)), _shares_of(quoted)
    qtext = str(quoted or "").lower()
    for v, groups in said.items():
        theirs = [t for t in (ev.get(v) or []) if t]
        if not theirs:
            continue
        for g in groups:
            if not g:
                continue
            mine = _restricting(g)
            fits = False
            for t in theirs:
                other = _restricting(t)
                if other - mine:
                    continue                       # (theirs narrower: this sentence widened it)
                if mine - other and not all(re.search(r"\b" + re.escape(x.split()[0][:6]), qtext) for x in mine - other):
                    continue                       # (this sentence narrower, and the papers never say so)
                fits = True
                break
            if not fits:
                return False
    return True


# ---- 6.6, round 110 (the owner, 21 Nov 2026 00:47: "Raise the app's validity."): a figure keeps when it was found. "Students slept
# 6.2 hours on school nights" where the paper's 6.2 hours were on weekends, or "in the rainy season" for its dry season, gives
# the figure to another time. Compared for the same figure, in its own clause, where both name a time.
_WHEN = (("weekend", re.compile(r"\bweek-?ends?\b|\bsaturdays?\b|\bsundays?\b", re.I)),
         ("weekday", re.compile(r"\bweek-?days?\b|\bschool[\s-]+(?:nights?|days?)\b|\bwork(?:ing)?[\s-]+days?\b|\bweeknights?\b", re.I)),
         ("dry", re.compile(r"\bdry[\s-]+seasons?\b|\bharmattan\b", re.I)),
         ("wet", re.compile(r"\b(?:rainy|wet|raining)[\s-]+seasons?\b|\brains\b", re.I)),
         ("dayshift", re.compile(r"\bday[\s-]+shifts?\b", re.I)),
         ("nightshift", re.compile(r"\bnight[\s-]+shifts?\b", re.I)))
_FIG_ANY = re.compile(r"(?<![\w.])(\d+(?:\.\d+)?)(?![\d])")


def _when_clauses(text):
    """{figure: [{times}]}: each figure and the times (weekends, school nights, the dry or rainy season, a shift) its clause names."""
    out = {}
    t = str(text or "")
    bounds = [0] + [m.end() for m in re.finditer(r"[;:]|\.(?!\d)|,\s*(?:but|while|whereas|and)\b|\b(?:but|whereas|while|compared\s+with)\b|"
                                                 r"\band\s+(?=\d)|,", t)] + [len(t)]
    lead = re.match(r"\s*(?:on|during|in|at|over|throughout)\s+(?:the\s+)?[^,;:.]{1,40},", t, re.I)       # ("On weekends, students slept…")
    lead_times = {name for name, rx in _WHEN if rx.search(lead.group(0))} if lead else set()
    for m in _FIG_ANY.finditer(t):
        a = max(b for b in bounds if b <= m.start())
        z = min(b for b in bounds if b > m.start())
        clause = t[a:z]
        times = {name for name, rx in _WHEN if rx.search(clause)} or set(lead_times)
        if times:
            try:
                out.setdefault(float(m.group(1)), []).append(times)
            except ValueError:
                pass
    return out


def when_ok(sentence, quoted):
    """The same figure said of another time than the papers' (school nights for their weekends, the rainy season for their
    dry season) is not theirs."""
    s = _when_clauses(bare(sentence))
    if not s:
        return True
    q = _when_clauses(quoted)
    for v, mine in s.items():
        if 1900 <= v <= 2100:
            continue
        theirs = q.get(v) or []
        if not theirs:
            continue
        for u in mine:
            if not any(u & t for t in theirs):
                return False
    return True


_REL = re.compile(r"\b(associat|link|relat|correlat|differ|predict)\w*", re.I)
_NEG = re.compile(r"\b(?:no|not|nor|neither|without|never|none|lack\s+of|absence\s+of|failed\s+to|did\s+not|does\s+not|do\s+not|"
                  r"was\s+not|were\s+not|is\s+not|are\s+not)\b|n['’]t\b", re.I)


# "sleep differed by group, but grades did not": the word said once, its "not" after it (6.6, round 23)
_IN_CLAUSE = r"(?:[^.;]|\.(?=\d))"            # (a decimal point is not the end of the sentence: "p = .02")
_ELIDED_NOT = re.compile(r"^" + _IN_CLAUSE + r"{0,80}?(?:,?\s*\bbut\b|;|,\s*whereas\b|,?\s*\bwhile\b)\s+(?:not\b|" + _IN_CLAUSE +
                         r"{0,50}?\b(?:did|does|do|was|were|is|are|had|has|have)\s*(?:not|n['’]t)\b"
                         r"(?!\s+(?:\w+\s+)?(?:associat|link|relat|correlat|differ|predict)))", re.I)


def _polarity(text):
    out = set()
    t = str(text or "")
    for m in _REL.finditer(t):
        before = re.split(r"[,;:]|\bbut\b|\bwhereas\b|\bwhile\b|\balthough\b", t[max(0, m.start() - 50):m.start()])[-1]
        b = " ".join(before.split()[-4:])
        b = re.sub(r"\bnot\s+(?:only|just|merely)\b|\bno\s+longer\b", "", b, flags=re.I)
        neg = bool(_NEG.search(b))
        out.add((m.group(1).lower()[:6], neg))
        if not neg and _ELIDED_NOT.search(t[m.end():]):
            out.add((m.group(1).lower()[:6], True))
    return out


def polarity_ok(sentence, quoted):
    """A link the papers found is not turned into none, nor none into a link: when the sentence says something is (or is
    not) associated, linked, related, correlated, different or predictive, and the papers' sentences say the same word only
    the other way round, the sentence turned the finding over."""
    s, q = _polarity(bare(sentence)), _polarity(quoted)
    for stem, neg in s:
        if (stem, not neg) in q and (stem, neg) not in q:
            return False
    return True


# ---- "all", "every", "none", "completely": only where the papers' sentences say as much
_ALL = re.compile(r"\b(all|every|everyone|everybody|none|nobody|no\s+one|entirely|completely|totally|wholly|universally|"
                  r"eliminat\w*|eradicat\w*)\b(?!\s+(?:of\s+)?(?:the\s+)?(?:above|same|time|in\s+all|along|but|too|over|around|right|"
                  r"the\s+more|the\s+less|year|day|week|night|morning))", re.I)
_ALL_EVID = {"all": r"\ball\b|\b100\s?%|\bevery\b|\beach\b", "every": r"\bevery\w*|\ball\b|\beach\b|\b100\s?%",
             "none": r"\bnone\b|\bno\b|\b0\s?%|\bnot\b|\bnobody\b|\bno\s+one\b"}


_ALL_IDIOM = re.compile(r"\b(?:most|first|best|worst|least)\s+of\s+all\b|\babove\s+all\b|\bafter\s+all\b|\ball\s+in\s+all\b|"
                        r"\bin\s+all\s*,|\bonce\s+and\s+for\s+all\b|\ball\s+the\s+(?:more|less|same)\b|\bnot\s+at\s+all\b", re.I)


def all_ok(sentence, quoted):
    """"All infections", "every student", "none of them", "completely", "eliminates": a claim about everyone or
    everything, only where the papers' sentences say as much ("all", "every", "100%", "none", "completely" there too).
    "Most of all", "above all", "after all", "in all," (in total) say nothing about everyone and are not counted."""
    q = str(quoted or "")
    s = _ALL_IDIOM.sub(" ", bare(sentence))
    # "nearly all", "almost all": a share of 85% or more, or the paper's own words (6.6, round 23)
    for m in _NEARLY_ALL.finditer(s):
        pcts = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)", q)]
        if not (re.search(r"\b(?:nearly|almost|virtually|practically)\s+(?:all|every)\b|\ball\b|\bevery\b", q, re.I)
                or any(v >= 85 for v in pcts)):
            return False
    s = _NEARLY_ALL.sub("nearly_all", s)
    for m in _ALL.finditer(s):
        w = re.sub(r"\s+", " ", m.group(1).lower())
        key = {"everyone": "every", "everybody": "every", "nobody": "none", "no one": "none"}.get(w, w)
        ev = _ALL_EVID.get(key)
        if ev and re.search(ev, q, re.I):
            continue
        if not ev and re.search(r"\b" + re.escape(w[:6]), q, re.I):
            continue
        return False
    return True


# ================================================================ 6.6, round 4: more claims held to the papers
# ---- "long-term", "lasting", "sustained": only where the papers followed people long enough (six months or more) or say so
_LONG_CLAIM = re.compile(r"\blong[\s-]?(?:term|lasting)\b|\blasting\s+(?:effects?|benefits?|improvements?|changes?|impacts?|gains?)\b|"
                         r"\b(?:sustained|persistent|enduring|durable)\s+(?:effects?|benefits?|improvements?|changes?|impacts?|gains?|"
                         r"reductions?|increases?)\b|\b(?:effects?|benefits?|gains?|improvements?)(?:\s+\w+){0,3}?\s+(?:lasted|persisted|endured|"
                         r"were\s+(?:sustained|maintained|kept\s+up))\b|\bin\s+the\s+long\s+(?:run|term)\b|\bover\s+(?:the\s+)?long\s+term\b|"
                         r"\bfor\s+years\s+(?:after|afterwards|later)\b", re.I)
_LONG_EVID = re.compile(r"\blong[\s-]?(?:term|lasting)\b|\bsustain(?:ed|s)?\b(?!\s+attention)|\bpersist\w*|\bmaintain\w*|\bendur\w*|"
                        r"\bdurab\w*|\byears?\s+later\b|\b(?:\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|"
                        r"fifteen|eighteen|twenty|twenty-four|thirty-six)[\s-]+(?:years?|months?)\b", re.I)


# (6.6, round 26) the follow-up found it gone: "no better six months later", "no longer", "faded", "returned to baseline"
_LONG_GONE = re.compile(r"\bno\s+(?:longer|better|different|difference)\b|\bnot\s+(?:sustained|maintained|persist\w*|last\w*|kept\s+up)\b|"
                        r"\b(?:did|does|do|was|were)\s+not\s+(?:persist|last|endure|remain|hold|continue|sustain|maintain)\w*|\bfaded\b|\bwaned\b|"
                        r"\bdisappeared\b|\bwore\s+off\b|\breturned\s+to\s+(?:baseline|(?:its|their|the)\s+(?:earlier|previous|former|original)\s+level)|"
                        r"\bshort[\s-]lived\b|\btemporary\b|\btransient\b", re.I)


def long_ok(sentence, quoted):
    """"Long-term benefits", "lasting effects", "the gains were sustained": only where the papers' sentences say so, or
    followed people for six months or more (ten weeks of a trial don't show a lasting effect). And never where that
    follow-up found the effect gone ("no better six months later", "faded", "returned to baseline"), 6.6 round 26."""
    if not _LONG_CLAIM.search(bare(sentence)):
        return True
    q = str(quoted or "")
    if _LONG_GONE.search(q) and not re.search(r"\b(?:sustain(?:ed|s)?|persist(?:ed|s)?|maintained|lasted|long[\s-]?(?:term|lasting))\b", q, re.I):
        return False
    for m in _LONG_EVID.finditer(q):
        w = m.group(0).lower()
        if "month" in w and not re.search(r"\b(?:year|long|sustain|persist|maintain|endur|durab)", w):
            n = re.match(r"(\d+|[a-z-]+)", w).group(1)
            n = int(n) if n.isdigit() else {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5}.get(n, 6)
            if n < 6:
                continue
        return True
    return False


# ---- "doubled", "twice as likely", "three times", "a two-fold rise", "halved": the papers' own ratio
_TIMES = re.compile(r"\b(?:(doubl\w*|twice(?!\s+(?:a|an|per|each|every|daily|weekly|monthly|yearly|over|as\s+often\s+as\s+(?:a|per))\b)|"
                    r"two-?fold|2-?fold)|(tripl\w*|three-?fold|3-?fold)|(quadrupl\w*|four-?fold|4-?fold)|(halv\w*|half\s+as\s+"
                    r"(?:likely|many|much|often|high|common))|(\d+(?:\.\d+)?|two|three|four|five|six|ten)\s*(?:times|-fold)\s+"
                    r"(?:as\s+|more|higher|greater|larger|likelier|less|lower|fewer|the|increase|risk|odds)\b)", re.I)
_TIMES_WORD = {"two": 2, "three": 3, "four": 4, "five": 5, "six": 6, "ten": 10}
_RATIO_STAT = re.compile(r"\b(?:a?ORs?|RRs?|HRs?|IRRs?|PRs?|odds\s+ratios?|risk\s+ratios?|hazard\s+ratios?|relative\s+risks?|"
                         r"prevalence\s+ratios?|rate\s+ratios?)\b\s*(?:\(\w+\)\s*)?(?:of|was|were|=|:|,)?\s*(\d+(?:\.\d+)?)", re.I)
_TIMES_FAMILY = {2: r"\bdoubl\w*|\btwice\b|\btwo-?fold\b|\b2-?fold\b|\b(?:2|two)(?:\.0)?\s*times\b|\bhalf\b|\bhalv\w*",
                 3: r"\btripl\w*|\bthree-?fold\b|\b3-?fold\b|\b(?:3|three)(?:\.0)?\s*times\b|\ba\s+third\b|\bone[\s-]third\b",
                 4: r"\bquadrupl\w*|\bfour-?fold\b|\b4-?fold\b|\b(?:4|four)(?:\.0)?\s*times\b|\ba\s+quarter\b|\bone[\s-]quarter\b"}


def _times_factor(m):
    if m.group(1):
        return 2.0
    if m.group(2):
        return 3.0
    if m.group(3):
        return 4.0
    if m.group(4):
        return 2.0                                  # halved: one half of the other, a ratio of two
    w = m.group(5).lower()
    return float(w) if w[0].isdigit() else float(_TIMES_WORD.get(w, 0))


def times_ok(sentence, quoted):
    """A ratio in words ("doubled", "twice as likely", "three times higher", "halved") is the papers' own: the same words,
    a ratio statistic (an odds, risk or hazard ratio) near it, or two of their figures that make it."""
    q = str(quoted or "")
    for m in _TIMES.finditer(bare(sentence)):
        f = _times_factor(m)
        if not f or f < 1.2:
            continue
        fam = _TIMES_FAMILY.get(int(f)) if float(f).is_integer() else None
        if fam and re.search(fam, q, re.I):
            continue
        if any(0.85 * f <= float(x) <= 1.15 * f for x in re.findall(r"(\d+(?:\.\d+)?)\s*(?:times|-?fold)\b", q, re.I)):
            continue                                    # "2.1 times higher" makes "twice as high"
        stats = [float(x) for x in _RATIO_STAT.findall(q)]
        if any(0.8 * f <= v <= 1.25 * f or (v > 0 and 0.8 * f <= 1 / v <= 1.25 * f) for v in stats):
            continue
        # two figures of one sentence of the papers that make the ratio (never a confidence level: "95% CI")
        pairs = False
        for one in re.split(r"(?<=[.!?;])\s+", q):
            nums = [float(x) for x in re.findall(r"(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)(?!\s*(?:CI|C\.I\.|confidence))", one)]
            if any(b > 0 and 0.85 * f <= a / b <= 1.15 * f for a in nums for b in nums if a != b):
                pairs = True
                break
        if pairs:
            continue
        return False
    return True


# ---- percent and percentage points: "rose by 10%" is not "rose by 10 percentage points" (from 20% to 30% is 50%)
_PCT_CHANGE = re.compile(r"\bby\s+(?:about\s+|around\s+|nearly\s+|almost\s+|over\s+|some\s+)?(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)"
                         r"|\b(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)\s+(?:increase|decrease|rise|fall|drop|reduction|"
                         r"gain|improvement|higher|lower|more|less|fewer|greater|smaller)\b", re.I)
_POINTS = re.compile(r"\b(\d+(?:\.\d+)?)\s+(?:percentage|%)[\s-]+points?\b|\b(\d+(?:\.\d+)?)\s?pp\b", re.I)
_FROM_TO = re.compile(r"\bfrom\s+(\d+(?:\.\d+)?)\s?%\s+(?:\w+\s+){0,3}?to\s+(\d+(?:\.\d+)?)\s?%", re.I)


def _near(a, b):
    return abs(a - b) <= max(0.5, 0.02 * abs(b))


def points_ok(sentence, quoted):
    """"Rose by 10%" where the papers say 10 percentage points (or from 20% to 30%, a rise of 50%) mixes up two different
    things, and so does "10 percentage points" where they say a rise of 10%."""
    s, q = bare(sentence), str(quoted or "")
    pts = [float(a or b) for a, b in _POINTS.findall(q)]
    spans = [(float(a), float(b)) for a, b in _FROM_TO.findall(q)]
    rel_said = [float(a or b) for a, b in _PCT_CHANGE.findall(q)]
    for a, b in _PCT_CHANGE.findall(s):
        v = float(a or b)
        if any(_near(v, x) for x in rel_said):
            continue                                       # the paper says the same "by 10%"
        if any(_near(v, x) for x in pts) or any(_near(v, abs(y - x)) and not _near(v, abs(y - x) / x * 100 if x else -1) for x, y in spans):
            return False                                   # it is a difference in points
    for a, b in _POINTS.findall(s):
        v = float(a or b)
        if any(_near(v, x) for x in pts) or any(_near(v, abs(y - x)) for x, y in spans):
            continue
        if any(_near(v, x) for x in rel_said):
            return False                                   # the paper's 10% is a relative change, not points
    return True


# ---- 6.6, round 12: the measure a figure is. Prevalence (who has it at a time) is not incidence (new cases over a time);
# mortality (deaths in a whole population) is not case fatality (deaths among those who have it); an odds ratio is not a
# risk ("twice the risk" from an odds ratio of 2 says more than the paper: when the outcome is common, odds run ahead of risk)
_MEASURE = {
    "prevalence": re.compile(r"\bprevalen(?:ce|t)\b|\bpoint\s+prevalence\b", re.I),
    "incidence": re.compile(r"\bincidence\b|\bincident\s+cases?\b|\bnew(?:ly\s+diagnosed)?\s+cases?\b|\bper\s+[\d,]+\s+(?:person|people|population|"
                            r"patient|child|woman|man|adult)s?[-\s](?:years?|months?)\b|\bperson[-\s]years?\b", re.I),
    "mortality": re.compile(r"\bmortality\b|\bdeath\s+rates?\b|\bdeaths?\s+per\s+[\d,]+\b", re.I),
    "fatality": re.compile(r"\bcase[-\s]fatality\b|\bCFR\b|\bfatality\s+rates?\b", re.I),
}
_PAIRS = (("incidence", "prevalence"), ("prevalence", "incidence"), ("mortality", "fatality"), ("fatality", "mortality"))
_RISK_SAID = re.compile(r"\b(?:twice|thrice|three\s+times|four\s+times|\d+(?:\.\d+)?[-\s]*(?:times|fold))\s+(?:the|a|as\s+high\s+a|higher|greater|increased)?\s*"
                        r"(?:risk|hazard)\b|\b(?:risk|hazard)\s+(?:was|were|is|are|of\s+\w+(?:\s+\w+)?\s+(?:was|were))\s+(?:doubled|tripled|halved|"
                        r"\d+(?:\.\d+)?\s*times)|\b(?:relative\s+risk|risk\s+ratio)\b|\b(?:doubled|tripled|halved)\s+the\s+risk\b", re.I)
_ODDS_EVID = re.compile(r"\bodds\s+ratios?\b|\b(?:a|adj\.?\s?|c)?OR\b|\bAOR\b|\bodds\b", re.I)
_RISK_EVID = re.compile(r"\b(?:risk|rate|hazard|prevalence|incidence)\s+ratios?\b|\brelative\s+risks?\b|\b(?:a)?(?:RR|HR|IRR|PR)\b|\bhazards?\b|"
                        r"\b(?:twice|\d+(?:\.\d+)?\s*times)\s+(?:the|a|as\s+high\s+a)?\s*risk\b", re.I)


# ---- 6.6, round 74: a figure keeps its own measure. The same number in the papers under another measure of its kind, and never
# under the sentence's own, is the paper's figure with another meaning: "a sensitivity of 85%" where the paper's 85% is its
# specificity; "a mean age of 21.4" where 21.4 is its median; a standard deviation that is its standard error; a hazard ratio
# that is its odds ratio. A sentence that lists measures and then their figures ("sensitivity and specificity were 85% and
# 90%", "… respectively") is not read for this: which figure goes with which is not in its words alone.
_LABEL_RX = (
    ("sensitivity", "dx", r"(?i:\bsensitivity\b)"), ("specificity", "dx", r"(?i:\bspecificity\b)"),
    ("ppv", "dx", r"(?i:\bpositive\s+predictive\s+values?\b)|\bPPVs?\b"), ("npv", "dx", r"(?i:\bnegative\s+predictive\s+values?\b)|\bNPVs?\b"),
    ("mean", "central", r"(?i:\bmean\b|\baverage\b)|(?<![A-Za-z])M(?=\s*=)"), ("median", "central", r"(?i:\bmedian\b)|\bMdn\b"),
    ("sd", "spread", r"(?i:\bstandard\s+deviations?\b)|\bSDs?\b"), ("se", "spread", r"(?i:\bstandard\s+errors?\b)|\bSEM?\b"),
    ("iqr", "spread", r"(?i:\binter-?quartile\s+range\b)|\bIQR\b"),
    ("or", "ratio", r"(?i:\bodds\s+ratios?\b)|\b(?:a|adj\.?\s?|c)?ORs?\b|\bAOR\b"),
    ("rr", "ratio", r"(?i:\b(?:risk|rate)\s+ratios?\b|\brelative\s+risks?\b)|\b(?:a)?RRs?\b|\bIRRs?\b"),
    ("hr", "ratio", r"(?i:\bhazard\s+ratios?\b)|\b(?:a)?HRs?\b"),
)
_LABEL_ANY = re.compile("|".join(f"(?P<{k}>{rx})" for k, _f, rx in _LABEL_RX))
_LABEL_FAMILY = {k: f for k, f, _ in _LABEL_RX}
_LABEL_FIG = re.compile(r"[^\d.;:]{0,40}?(?<![\w.])(\d+(?:\.\d+)?)(?!\d)")
_LABEL_BREAK = re.compile(r"\b(?:but|while|whereas|although|compared|than|versus|vs)\b", re.I)     # (another clause: not its figure)


def _labelled(text):
    """[(measure, its figure as written)] of each sentence that names its measures one by one, each with its own figure."""
    out = []
    for snt in re.split(r"(?<=[.!?;])\s+", str(text or "")):
        if re.search(r"\brespectively\b", snt, re.I):
            continue
        marks = list(_LABEL_ANY.finditer(snt))
        listed = any(_LABEL_FAMILY[a.lastgroup] == _LABEL_FAMILY[b.lastgroup] and not re.search(r"\d", snt[a.end():b.start()])
                     for a, b in zip(marks, marks[1:]))
        if listed:
            continue                                # ("sensitivity, specificity and PPV were 85%, 90% and 70%")
        for m in marks:
            f = _LABEL_FIG.match(snt, m.end())
            if f and not _LABEL_ANY.search(snt[m.end():f.start(1)]) and not _LABEL_BREAK.search(snt[m.end():f.start(1)]):
                out.append((m.lastgroup, f.group(1)))
    return out


def _fig_same(said, theirs):
    """The sentence's figure is the paper's, or the paper's rounded to fewer decimals (21 for 21.4)."""
    try:
        a, b = float(said), float(theirs)
    except ValueError:
        return False
    if abs(a - b) < 1e-9:
        return True
    da = len(said.split(".", 1)[1]) if "." in said else 0
    db = len(theirs.split(".", 1)[1]) if "." in theirs else 0
    return db > da and abs(math.floor(b * 10 ** da + 0.5) / 10 ** da - a) < 1e-9


def label_ok(sentence, quoted):
    """Each figure keeps its measure (6.6, round 74): not one the papers give, the same number, under another measure of its
    kind and never under the sentence's own."""
    said = _labelled(bare(sentence))
    if not said:
        return True
    ev = _labelled(quoted)
    for k, v in said:
        if any(k2 == k and _fig_same(v, v2) for k2, v2 in ev):
            continue
        if any(k2 != k and _LABEL_FAMILY[k2] == _LABEL_FAMILY[k] and _fig_same(v, v2) for k2, v2 in ev):
            return False
    return True


def measure_ok(sentence, quoted):
    """A figure keeps its measure: "incidence" only where the papers measured new cases (not "prevalence"), and the other
    way round; "mortality" not where they gave case fatality, and the other way round; "twice the risk", "relative risk"
    or "the risk was doubled" not where the papers give an odds ratio (and no risk, rate or hazard ratio). 6.6, round 74:
    a sensitivity, a mean, a standard deviation, a ratio is not the paper's specificity, median, standard error or another
    ratio of the same number (label_ok)."""
    s, q = bare(sentence), str(quoted or "")
    for said, other in _PAIRS:
        if _MEASURE[said].search(s) and not _MEASURE[said].search(q) and _MEASURE[other].search(q):
            return False
    if _RISK_SAID.search(s) and _ODDS_EVID.search(q) and not _RISK_EVID.search(q):
        return False
    return label_ok(sentence, quoted)


# ---- 6.6, round 19: a share keeps its people ("38% of late users" is not "38% of students")
_PCT_N = r"(\d+(?:\.\d+)?)\s?(?:%|percent\b|per\s+cent\b)"
_PCT_OF_RX = re.compile(_PCT_N + r"\s+of\s+(?:the\s+|all\s+|these\s+|those\s+|their\s+)?([^.;:()\d]{1,80})", re.I)
_AMONG_RX = re.compile(r"\bamong\s+(?:the\s+|all\s+)?([^.;:(),]{1,60}?)\s*,?\s+" + _PCT_N + r"|"
                       r"(?:^|[.;:]\s+)of\s+(?:the\s+|all\s+)?([^.;:(),]{1,60}?)\s*,\s*" + _PCT_N, re.I)   # (round 111: "Of the pregnant women …, 38%")
_DEN_STOP = {"and", "or", "but", "were", "was", "who", "which", "that", "in", "at", "on", "with", "had", "have", "has", "reported", "slept",
             "said", "did", "do", "are", "is", "from", "by", "for", "to", "used", "showed", "felt", "than", "compared", "while", "whereas",
             "aged", "across", "during", "over", "after", "before", "per", "each", "would", "could", "might", "may", "will"}
_DEN_GENERIC = {"participant", "respondent", "subject", "sample", "people", "individual", "person", "population", "interviewee", "total",
                "them", "all", "case", "those"}


def _den_words(phrase):
    """The people a share is of, as words ("late users", "first-year students"): up to four words, stopping at a verb or a
    small word; each without its plural s."""
    out = []
    for w in re.findall(r"[a-z][\w-]*", str(phrase or "").lower()):
        if w in _DEN_STOP or len(out) >= 4:
            break
        out.append(re.sub(r"(?<=[a-z]{3})(?:es|s)$", "", w))
    return [w for w in out if len(w) >= 3 and w not in ("the", "all", "other", "these", "those")]


def _shares_of(text):
    """{figure: [the words of whom it is a share]} for "N% of X" and "among X, N%"."""
    out = {}
    for m in _PCT_OF_RX.finditer(str(text or "")):
        out.setdefault(float(m.group(1)), []).append(_den_words(m.group(2)))
    for m in _AMONG_RX.finditer(str(text or "")):      # (6.6, round 111: "Of the pregnant women …, 38%" too)
        g = [x for x in m.groups() if x is not None]
        try:
            out.setdefault(float(g[1]), []).append(_den_words(g[0]))
        except (ValueError, IndexError):
            continue
    return out


def denominator_ok(sentence, quoted):
    """A share keeps its people: where the sentence gives "N% of X" (or "among X, N%") and the papers give the same N% of
    other people ("38% of late users"), the sentence says what the papers don't. People named only generally
    ("participants", "respondents") on either side are taken to be the same."""
    said, ev = _shares_of(bare(sentence)), _shares_of(quoted)
    for v, groups in said.items():
        theirs = ev.get(v)
        if not theirs:
            continue                           # (the papers give that figure some other way: not compared here)
        for g in groups:
            if not g or set(g) & _DEN_GENERIC:
                continue
            if any(not t or set(t) & _DEN_GENERIC or set(g) & set(t) for t in theirs):
                continue
            return False
    return True


# ---- 6.6, round 20: a figure keeps its period ("lifetime" is not "in the past year"; "in the past month" is not "in the past week")
_PERIODS = (
    ("lifetime", re.compile(r"\blife-?time\b|\bever\s+(?:had|used|been|experienced|smoked|drunk|tried)\b|\bat\s+some\s+point\s+in\s+(?:their|his|her)\s+lives?\b", re.I)),
    ("year", re.compile(r"\b(?:past|last|previous|preceding)[\s-]+(?:year|12\s+months|twelve\s+months)\b|\b(?:12|twelve)[\s-]+months?\s+prevalence\b|\bpast-year\b|"
                        r"\bannual(?:ly)?\s+prevalence\b|\bwithin\s+the\s+(?:past|last)\s+(?:year|12\s+months)\b", re.I)),
    ("month", re.compile(r"\b(?:past|last|previous|preceding)[\s-]+(?:month|30\s+days|thirty\s+days|four\s+weeks|4\s+weeks)\b|\b(?:30|thirty)[\s-]+days?\s+prevalence\b|\bpast-month\b", re.I)),
    ("2weeks", re.compile(r"\b(?:past|last|previous|preceding)[\s-]+(?:two|2)\s+weeks\b|\b(?:14|fourteen)\s+days\b", re.I)),
    ("week", re.compile(r"\b(?:past|last|previous|preceding)[\s-]+(?:week|7\s+days|seven\s+days)\b|\bpast-week\b", re.I)),
)


def _periods(text):
    return {name for name, rx in _PERIODS if rx.search(str(text or ""))}


def period_ok(sentence, quoted):
    """A figure keeps its period: a sentence that says "lifetime", "in the past year", "in the past month" or "in the past
    week" where the papers give another of these (and not that one) says what they don't."""
    said, theirs = _periods(bare(sentence)), _periods(quoted)
    return not (said and theirs and not (said & theirs))


# ---- a dose-response ("the more…, the more…", "each additional hour", "in a graded way"): only where the papers show a gradient
_DOSE_CLAIM = re.compile(r"\bdose[\s-]+(?:response|dependent)\b|\bthe\s+(?:more|higher|greater|longer|less|lower|fewer|later)\b"
                         r"[^.;]{1,80}?,?\s+the\s+(?:more|higher|greater|longer|less|lower|fewer|worse|better|poorer|later|shorter)\b|"
                         r"\b(?:each|every)\s+(?:additional|extra|further)\s+\w+|\bfor\s+(?:each|every)\s+(?:additional\s+|extra\s+)?"
                         r"(?:hour|year|unit|point|day|kilogram|kg|cigarette|drink|session|minute|week|month|serving)\b|"
                         r"\bin\s+a\s+(?:graded|dose-dependent|step-?wise)\s+(?:manner|fashion|way)\b|\bgraded\s+(?:association|"
                         r"relationship|link)\b", re.I)
_DOSE_EVID = re.compile(r"\bdose[\s-]+(?:response|dependent)\b|\bthe\s+(?:more|higher|greater|longer|less|lower|fewer|later)\b|"
                        r"\b(?:each|every|per)\s+(?:additional|extra|one[\s-]|1[\s-])?\s*(?:hour|year|unit|point|day|kilogram|kg|"
                        r"cigarette|drink|session|minute|week|month|serving|standard\s+deviation|sd)\b|\bgraded\b|\bp[\s-]?(?:for[\s-])?"
                        r"trend\b|\b(?:linear|monotonic|step-?wise)\s+(?:trend|association|relationship|increase|decrease)\b|\bgradient\b|"
                        r"\bincreas\w+\s+(?:steadily\s+)?with\s+(?:each|every|increasing|more|higher|longer|greater)\b|"
                        r"\bdecreas\w+\s+(?:steadily\s+)?with\s+(?:each|every|increasing|more|higher|longer|greater)\b", re.I)


def dose_ok(sentence, quoted):
    """"The more time on social media, the less sleep", "each additional hour", "in a graded way": a gradient, which a
    comparison of two groups doesn't show; only where the papers' sentences show one."""
    return not _DOSE_CLAIM.search(bare(sentence)) or bool(_DOSE_EVID.search(str(quoted or "")))


# ================================================================ 6.6, round 6: groups of countries, agreement, "currently"
# ---- "developing countries", "low- and middle-income countries", "the Global South": what the papers say they studied
_REGION_NOUN = r"(?:countr(?:y|ies)|nations?|settings?|economies|economy|regions?|world)"
_REGION_GROUP = re.compile(r"\b(?:developing|developed|industriali[sz]ed|western|high-?income|low-?income|middle-?income|lower-?middle-?income|"
                           r"upper-?middle-?income|resource-(?:limited|poor|constrained)|low-?resource)\s+" + _REGION_NOUN + r"\b|"
                           r"\blow-?\s+and\s+middle-?income\s+" + _REGION_NOUN + r"\b|\bLMICs?\b|\bHICs?\b|\bthe\s+global\s+(?:south|north)\b|"
                           r"\bthird[\s-]world\b", re.I)
_REGION_FAMILY = (("low", re.compile(r"\b(?:developing|low-?income|middle-?income|lower-?middle-?income|resource-(?:limited|poor|constrained)|"
                                     r"low-?resource)\s+" + _REGION_NOUN + r"\b|\blow-?\s+and\s+middle-?income\b|\bLMICs?\b|\bglobal\s+south\b|"
                                     r"\bthird[\s-]world\b", re.I)),
                  ("high", re.compile(r"\b(?:developed|industriali[sz]ed|western|high-?income|upper-?middle-?income)\s+" + _REGION_NOUN +
                                      r"\b|\bHICs?\b|\bglobal\s+north\b", re.I)))


_REGION_MANY = re.compile(r"countries|nations|settings|economies|regions|world|LMICs|HICs|\blow-?\s+and\s+middle|global\s+(?:south|north)|"
                          r"third[\s-]world", re.I)


def region_ok(sentence, evidence):
    """"Developing countries", "low- and middle-income countries", "the Global South", "high-income countries": a group of
    countries is the papers' own words (one country's students are not "students in developing countries"). And a group
    is not one of them (6.6, round 23): a paper that calls its one country "a lower-middle-income country" backs "in
    Ghana, a lower-middle-income country", not "in developing countries"."""
    ev = str(evidence or "")
    for m in _REGION_GROUP.finditer(bare(sentence)):
        rx = next((r for _k, r in _REGION_FAMILY if r.search(m.group(0))), None)
        if rx is None or not rx.search(ev):
            return False
        if _REGION_MANY.search(m.group(0)) and not any(_REGION_MANY.search(e.group(0)) for e in rx.finditer(ev)):
            return False
    return True


# ---- studies agree, studies disagree: what the papers show of each other
_AGREE = re.compile(r"\b(?:consistent(?:ly)?\s+(?:shown|found|reported|demonstrated|observed|evidence|findings|results)|"
                    r"(?:studies|research|evidence|findings|authors|researchers|scholars)\s+(?:\w+\s+){0,2}?(?:agree|concur)|consensus|"
                    r"consistently|unanimous\w*|well[\s-]established|it\s+is\s+(?:well\s+)?(?:known|established|accepted))\b", re.I)
_DISAGREE = re.compile(r"\b(?:disagree\w*|conflicting|contradict\w*|inconsistent|mixed\s+(?:results|findings|evidence)|"
                       r"(?:results|findings|studies)\s+(?:\w+\s+){0,2}?(?:differ|vary|varied|are\s+mixed|were\s+mixed)|contrasting\s+"
                       r"(?:results|findings)|debated|controversial|no\s+consensus|(?:are|is|were|was|remains?)\s+(?:still\s+)?mixed)\b", re.I)
_SAYS_AGREE = re.compile(r"\bconsisten\w*|\bconsensus\b|\bagree\w*|\bestablished\b|\bwell[\s-]known\b|\bconfirm\w*|\breplicat\w*", re.I)
_SAYS_DISAGREE = re.compile(r"\b(?:disagree\w*|conflicting|contradict\w*|inconsisten\w*|mixed|differ\w*|vari\w*|debate\w*|controvers\w*|"
                            r"unclear|no\s+consensus|heterogene\w*|equivocal)\b", re.I)
_FINDING_WORDS = re.compile(r"\b(?:improv|increas|reduc|decreas|rais|lower|higher|associat|link|effect|predict|correlat)\w*", re.I)


def agree_ok(sentence, n_papers, quoted):
    """"Studies consistently show", "it is well established": two papers or more, or a paper that says as much.
    "Studies disagree", "the findings are mixed", "conflicting evidence": papers that do differ (one found a link, another
    found none), or a paper that says the evidence is mixed."""
    s, q = bare(sentence), str(quoted or "")
    if _AGREE.search(s) and (n_papers or 0) < 2 and not _SAYS_AGREE.search(q):
        return False
    if _DISAGREE.search(s) and not _SAYS_DISAGREE.search(q):
        if (n_papers or 0) < 2:
            return False
        pol = _polarity(q)
        if not (pol and any(neg for _s, neg in pol) and any(not neg for _s, neg in pol)) and not (_NEG.search(q) and _FINDING_WORDS.search(q)):
            return False
    return True


# ---- "currently", "today", "now", with a figure: only from papers of the last five years
_CURRENT = re.compile(r"\b(?:currently|at\s+present|presently|nowadays|these\s+days|today|now|at\s+the\s+moment|in\s+the\s+present\s+day)\b", re.I)


_FIGURE_SAID = re.compile(r"\d|\b(?:one|two|three|four|five|six|seven|eight|nine)\s+(?:in|out\s+of)\s+(?:\d+|two|three|four|five|six|"
                          r"seven|eight|nine|ten|twenty|a\s+hundred)\b|\bhalf\b|\b(?:thirds?|quarters?|fifths?|tenths?)\b|"
                          r"\b(?:the\s+|a\s+)?(?:majority|minority)\b|\bmost\b", re.I)


def current_ok(sentence, pub_years=(), this_year=None):
    """"Currently, 38% of students…", "today, one in three adults…": a figure said of today stands only on papers from
    the last five years."""
    s = bare(sentence)
    if not this_year or not _CURRENT.search(s) or not _FIGURE_SAID.search(s):
        return True
    ys = [int(y) for y in (pub_years or ()) if str(y).isdigit()]
    return not ys or max(ys) >= this_year - RECENT_YEARS


def claim_problems(sentence, quoted, papers_text, n_papers, pub_years=(), all_text=None, this_year=None):
    """What a sentence says that its papers don't (empty when nothing): quoted = the papers' sentences that back it
    (the checker named them), papers_text = what those papers say of their OWN study (their titles and every sentence
    that is not their background: where, when and whom they studied), all_text = everything of them that was read (for
    words in quotation marks; papers_text when not given), n_papers = how many different papers it cites, pub_years =
    their years. A place, the world, a group or a year that only a paper's background names is not where, when or whom
    it studied: it backs the sentence only when the checker named that sentence (it is then in quoted)."""
    out = []
    ev = " ".join((quoted or "", papers_text or ""))
    if not years_ok(sentence, ev, pub_years):
        out.append("year")
    if not places_ok(sentence, ev):
        out.append("place")
    if not world_ok(sentence, ev):
        out.append("world")
    if not many_ok(sentence, n_papers, ev):
        out.append("many")
    if not quotes_ok(sentence, " ".join((quoted or "", all_text or papers_text or ""))):
        out.append("quote")
    if not cause_ok(sentence, quoted or papers_text):
        out.append("cause")
    if not groups_ok(sentence, ev) or not qualifiers_ok(sentence, ev):
        out.append("group")
    # 6.5, further: the kind of study, "significant", certainty, superlatives, shares, the direction of a finding
    if not design_ok(sentence, ev):
        out.append("design")
    if not significance_ok(sentence, quoted or papers_text):
        out.append("significance")
    if not certainty_ok(sentence, quoted or papers_text):
        out.append("certainty")
    if not superlative_ok(sentence, quoted or papers_text):
        out.append("superlative")
    if not shares_ok(sentence, quoted or papers_text):
        out.append("share")
    if not direction_ok(sentence, quoted or papers_text):
        out.append("direction")
    # 6.5, higher: a trend, a study's scale, a share in words, more words of cause, a hedge, "recent", "the first", polarity
    if not trend_ok(sentence, quoted or papers_text):
        out.append("trend")
    if not scale_ok(sentence, ev):
        out.append("scale")
    if not ratio_ok(sentence, quoted or papers_text):
        out.append("share")
    if not cause2_ok(sentence, quoted or papers_text) and "cause" not in out:
        out.append("cause")
    if not hedge_ok(sentence, quoted or papers_text):
        out.append("hedge")
    if not recent_ok(sentence, pub_years, this_year):
        out.append("recent")
    if not first_ok(sentence, quoted or papers_text):
        out.append("first")
    if not polarity_ok(sentence, quoted or papers_text) and "direction" not in out:
        out.append("direction")
    # 6.6, round 101: a level ("knowledge was poor") or a share's "not" ("38% did not …") said the other way; round 102: the
    # higher side of a comparison (girls and boys, intervention and control…), a link's sign, a risk factor or protective
    if "direction" not in out and not (level_ok(sentence, quoted or papers_text) and share_sign_ok(sentence, quoted or papers_text)
                                       and comparison_ok(sentence, quoted or papers_text) and sign_ok(sentence, quoted or papers_text)):
        out.append("direction")
    # 6.6, round 103: "the most common", "the leading", "the strongest…" given to what the papers give it to
    if "superlative" not in out and not superlative_whose_ok(sentence, quoted or papers_text):
        out.append("whose")
    # 6.6, round 106: a figure from one place said of a whole country
    if "scale" not in out and not scope_ok(sentence, ev):
        out.append("scale")
    if not all_ok(_NEARLY_ALL.sub("nearly_all", sentence) if "share" in out else sentence, quoted or papers_text):
        out.append("all")                     # ("nearly all nurses" for 64% is said once: "share")
    # 6.6, round 4: a lasting effect, a ratio in words, percent and percentage points, a dose-response
    if not long_ok(sentence, quoted or papers_text):
        out.append("long")
    if not times_ok(sentence, quoted or papers_text):
        out.append("times")
    if not points_ok(sentence, quoted or papers_text):
        out.append("points")
    if not dose_ok(sentence, quoted or papers_text):
        out.append("dose")
    # 6.6, round 6: a group of countries, studies agreeing or disagreeing, "currently" with a figure
    if not region_ok(sentence, ev):
        out.append("region")
    if "many" not in out and not agree_ok(sentence, n_papers, quoted or papers_text):
        out.append("agree")                   # ("studies agree" from one paper is said once: "many")
    if not current_ok(sentence, pub_years, this_year):
        out.append("current")
    # 6.6, round 12: prevalence and incidence, mortality and case fatality, odds and risk
    if not measure_ok(sentence, quoted or papers_text):
        out.append("measure")
    # 6.6, round 19: a share keeps its people; round 109: and the words that narrow them ("pregnant women" is not "women")
    if not denominator_ok(sentence, quoted or papers_text) or not narrowing_ok(sentence, quoted or papers_text):
        out.append("whom")
    # 6.6, round 108: how often ("daily" is not "weekly"); round 110: when (school nights are not weekends)
    if not frequency_ok(sentence, quoted or papers_text):
        out.append("often")
    if not when_ok(sentence, quoted or papers_text):
        out.append("when")
    # 6.6, round 20: a figure keeps its period (lifetime, the past year, month or week)
    if not period_ok(sentence, quoted or papers_text):
        out.append("period")
    # 6.6, round 21: what research has not done, only where a paper says so
    if not absence_ok(sentence, quoted or papers_text):
        out.append("absence")
    # 6.6, round 22: a correlation's strength in words keeps to its size
    if not strength_ok(sentence, quoted or papers_text):
        out.append("strength")
    # 6.6, round 60: how strong the evidence is ("it is well established" from one paper is said once: "many")
    if "many" not in out and not evidence_ok(sentence, quoted or papers_text, n_papers):
        out.append("evidence")
    # 6.6, round 5: a cause, where the papers' whole text only observed (a survey, a link) and never speaks of a cause, an
    # effect or a trial (the quoted sentence alone, "slept 48 minutes less", doesn't say which)
    if "cause" not in out and (_CAUSE_CLAIM.search(bare(sentence)) or _unhedged(_CAUSE_CLAIM2, bare(sentence))):
        whole = " ".join((quoted or "", papers_text or "", all_text or ""))
        if not _CAUSE_EVID.search(whole) and _OBSERVED.search(whole):
            out.append("cause")
    # 6.6, round 8: the sentence's own figures disagree (a p that is not what its statistic gives, an estimate outside its
    # interval, a percentage that is not its count of its total): nothing sourced may say it, even as its paper says it
    if _st66.problems(sentence):
        out.append("stats")
    return out


_OBSERVED = re.compile(r"\b(?:survey\w*|cross-?sectional|questionnaires?|interview\w*|observational|cohorts?|case-?control|"
                       r"associat\w*|correlat\w*|linked|relationships?)\b", re.I)
_NEARLY_ALL = re.compile(r"\b(?:nearly|almost|virtually|practically)\s+(?:all|every)\b", re.I)


# The rules every checker is given (Research, research writing, manual statements), in the AI's own words
STRENGTH_RULES = (
    "A paper backs a sentence fully only when it says the same thing with the same strength: a link, an association or a "
    "correlation is not a cause; \"may\", \"suggests\" or \"is linked to\" is not \"does\"; what one group showed (nurses, "
    "adolescents, women) is not shown for another group or for everyone; what one place showed (one town, one country) is "
    "not shown for another place, a region or the world; one study is not \"studies\" or \"research\", and two are not \"several\", "
    "\"many\" or \"most\" studies nor \"a growing body of research\"; a year, a figure or a "
    "place the sentence gives must be the paper's own; the kind of study (a trial, a cohort, a survey, a review) is the "
    "paper's; a finding keeps its direction; \"significant\", \"proves\", \"always\", \"the most\" or \"most students\" "
    "only where the paper says as much; a change over time (\"rising\", \"has increased\", \"increasingly\") only where the "
    "paper shows one; \"nationally representative\", \"large-scale\" or \"multicentre\" only where the paper says so; a share "
    "in words (\"one in three\", \"two-thirds\") is the paper's own figure; \"worsens\", \"affects\", \"improves\" or "
    "\"has an impact on\" are words of cause; \"may contribute\" is not \"contributes\"; \"recent\" means the last five years, "
    "\"the first\" only where the paper says it; \"no link\" is not \"a link\"; \"long-term\", \"lasting\" or \"sustained\" only "
    "where the paper followed people for six months or more or says so; \"doubled\", \"twice as likely\" or \"three times\" "
    "only where the paper's own figures or ratio make it; \"by 10%\" is not \"by 10 percentage points\" (from 20% to 30% is a "
    "rise of 10 points, or of 50%); \"the more…, the more…\" or \"each additional hour\" only where the paper shows a "
    "gradient; \"developing countries\", \"low- and middle-income countries\" or \"the Global South\" only where the paper "
    "studied them, not one country of them; \"studies agree\" or \"consistently\" needs two papers or a paper that says so, "
    "\"studies disagree\" or \"mixed findings\" needs papers that do differ or a paper that says so; \"currently\" or "
    "\"today\" with a figure only from a paper of the last five years; prevalence (who has it) is not incidence (new cases "
    "over a time), mortality (deaths in a population) is not case fatality (deaths among those who have it), and an odds "
    "ratio is not a risk (\"twice the odds\", not \"twice the risk\"); a figure keeps its own measure (a sensitivity is not "
    "the specificity, a mean not the median, a standard deviation not the standard error, an odds ratio not a hazard ratio); "
    "a share keeps its people (\"38% of late users\" is not "
    "\"38% of students\"); a figure keeps its period (lifetime is not the past year, the past month is not the past week); "
    "\"no study has…\", \"little is known\", \"few studies\" or \"a gap in the literature\" only where the paper itself says so "
    "(\"no study has\" only where it says there is none, as \"the first study\" does); secondary-school students are not "
    "university students, first-year students are not final-year students, rural is not urban, private schools are not public schools; "
    "\"a strong\", \"moderate\" or \"weak\" correlation keeps to the size the paper gives (r = .23 is weak) and to the "
    "paper's own word for it (\"moderate\" is not \"strong\"), and \"a small\", \"medium\" or \"large\" effect to its "
    "d or η² (d = 0.25 is small); a level keeps the paper's word for it (knowledge it calls poor is not good, low "
    "uptake is not high) and a share its \"not\" (38% who did not is not 38% who did; \"unvaccinated\" is not \"vaccinated\"); "
    "a comparison keeps its sides (girls who slept more than boys are not boys who slept more than girls; the intervention "
    "group is not the control group), a link its sign (\"positively\" is not \"negatively\"; a risk factor is not "
    "protective) and \"the most common\" or \"the leading\" what the paper gives it to; a figure from one college, hospital, "
    "district or city is that place's, never the whole country's, unless the paper is national; a share keeps the words that "
    "narrow its people (38% of women is not 38% of pregnant women), its \"how often\" (daily is not weekly) and its "
    "time (school nights are not weekends; the dry season is not the rainy season); "
    "a sum of money keeps its currency (US$50 is not GH₵50); a figure's year is the year its data are from as the paper "
    "gives it, not the year the paper came out; \"not significant\" is not \"significant\", nor \"improved\" \"worsened\", "
    "nor \"later\" \"earlier\"; one country the paper calls \"a lower-middle-income country\" is not \"developing countries\"; "
    "\"strong\", \"robust\" or \"compelling\" evidence, \"the evidence is clear\", \"firmly established\", \"clearly shows\" or "
    "\"definitely\" only where the paper rates its evidence so; \"it is known that\" needs three papers, as \"well established\" does. "
    "A sentence that goes further than its paper in any of these ways is not backed by it.")


# ---------------------------------------------------------------- sources
def _title_words(t):
    return set(re.findall(r"[a-z0-9]{3,}", re.sub(r"<[^>]+>", " ", str(t or "")).lower()))


def same_title(a, b):
    """Two titles are one paper's: most of their words shared (a subtitle or a typographic difference doesn't matter)."""
    wa, wb = _title_words(a), _title_words(b)
    if not wa or not wb:
        return False
    small = min(len(wa), len(wb))
    return len(wa & wb) >= max(2, round(0.7 * small))


def _cr_year(msg):
    for k in ("published-print", "published-online", "issued", "published", "created"):
        try:
            y = int(((msg.get(k) or {}).get("date-parts") or [[None]])[0][0])
            if 1500 <= y <= 2100:
                return y
        except (TypeError, ValueError, IndexError):
            continue
    return None


_CR_CHAPTER = ("book-chapter", "book-section", "book-part")
_CR_BOOKISH = ("book", "monograph", "edited-book", "reference-book", "book-set", "book-series") + _CR_CHAPTER


def journal_abbrev(msg):
    """A journal article's journal as its publisher abbreviated it (Crossref's short-container-title), written as NLM writes
    abbreviations: no full stops ("J. Sleep Res." -> "J Sleep Res"). "" when there is none, or when it is the full title
    itself, or not a journal article."""
    if str(msg.get("type") or "") != "journal-article":
        return ""
    full = " ".join(str(x) for x in (msg.get("container-title") or [])[:1]).strip()
    short = " ".join(str(x) for x in (msg.get("short-container-title") or [])[:1]).strip()
    short = re.sub(r"\s+", " ", short.replace(".", " ")).strip(" ,;:")
    if not short or not full or short.lower() == re.sub(r"\s+", " ", full.replace(".", " ")).strip().lower():
        return ""
    if len(short) > len(full) or not re.search(r"[A-Za-z]", short) or len(short) < 3:
        return ""
    return short[:120]


def _cr_authors(msg, field="author"):
    out = []
    for a in msg.get(field) or []:
        if not isinstance(a, dict):
            continue
        fam, giv, name = str(a.get("family") or "").strip(), str(a.get("given") or "").strip(), str(a.get("name") or "").strip()
        if fam:
            out.append((giv + " " + fam).strip())
        elif name:
            out.append(name)
    return out


def _cr_pages(msg):
    page = str(msg.get("page") or "").strip()
    art = str(msg.get("article-number") or "").strip()
    if page:
        first, _, last = page.partition("-")
        return {"first_page": first.strip(), "last_page": last.strip()} if last.strip() else {"first_page": first.strip()}
    if art:
        return {"first_page": art, "article": True}
    return {}


_CR_RETRACTED = {"retraction", "withdrawal", "removal", "partial_retraction"}
_CR_CONCERN = {"expression_of_concern", "expression-of-concern"}          # 6.5, further: left out too, and said
_CR_CORRECTED = {"correction", "erratum", "corrigendum"}                  # 6.6, round 94: kept, and said


def _cr_corrected(msg):
    """6.6, round 94: the year of the latest correction its journal published to a paper (its registry's "updated-by": a
    correction, an erratum or a corrigendum); 0 when the record gives no year; None when there is none. A corrected paper
    is still cited, and the person is told: its abstract (what RamzAI checked against) may be from before the correction."""
    best = None
    for u in ((msg or {}).get("updated-by") or []) if isinstance(msg, dict) else []:
        if not isinstance(u, dict) or str(u.get("type") or "").lower() not in _CR_CORRECTED:
            continue
        try:
            y = int(((u.get("updated") or {}).get("date-parts") or [[None]])[0][0])
        except (TypeError, ValueError, IndexError):
            y = 0
        best = max(best or 0, y if 1900 <= y <= 2100 else 0)
    return best


# ---- 6.6, round 3: a preprint that has been published since is cited as the published paper
def _doi(x):
    d = re.sub(r"^\s*(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", "", str(x or ""), flags=re.I).strip().lower()
    return d if re.fullmatch(r"10\.\d{4,9}/\S{1,200}", d) else ""


def _cr_preprint_of(msg):
    """The DOI of the paper a preprint was published as (its registry's "is-preprint-of"), or ""."""
    if not isinstance(msg, dict) or (str(msg.get("type") or "") != "posted-content" and str(msg.get("subtype") or "") != "preprint"):
        return ""
    for rel in ((msg.get("relation") or {}).get("is-preprint-of") or []):
        if isinstance(rel, dict) and str(rel.get("id-type") or "").lower() == "doi" and _doi(rel.get("id")):
            return _doi(rel.get("id"))
    return ""


def _cr_withdrawn(msg):
    ups = [str((u or {}).get("type") or "").lower() for u in ((msg or {}).get("updated-by") or [])]
    if any(u in _CR_RETRACTED for u in ups) or str((msg or {}).get("subtype") or "").lower() in _CR_RETRACTED:
        return "retracted"
    if any(u in _CR_CONCERN for u in ups):
        return "concern"
    return ""


def published_abstract(msg, work):
    """The published paper's own abstract: its registry's (the publisher's deposit), else OpenAlex's; "" when neither
    gives a real one (200 characters or more)."""
    import html as _html
    a = ""
    raw = str((msg or {}).get("abstract") or "")
    if raw:
        a = re.sub(r"\s+", " ", _html.unescape(re.sub(r"<[^>]+>", " ", raw))).strip()
        a = re.sub(r"^(?:abstract|summary)\b[\s:.\-–]*", "", a, flags=re.I).strip()
    inv = (work or {}).get("abstract_inverted_index") if isinstance(work, dict) else None
    if len(a) < 200 and isinstance(inv, dict):
        pos = {}
        for word, idx in inv.items():
            for i in (idx or []):
                if isinstance(i, int) and 0 <= i < 6000:
                    pos[i] = str(word)
        b = " ".join(pos[i] for i in sorted(pos))
        a = b if len(b) > len(a) else a
    return a if len(a) >= 200 else ""


def published_copy(work):
    """A free copy of the published paper itself (its published or accepted version, never the preprint again):
    (pdf, landing page) or (None, None)."""
    loc = (work or {}).get("best_oa_location") if isinstance(work, dict) else None
    loc = loc if isinstance(loc, dict) else {}
    pdf = str(loc.get("pdf_url") or "")
    if str(loc.get("version") or "") in ("publishedVersion", "acceptedVersion") and pdf.startswith("https://"):
        return pdf[:600], (str(loc.get("landing_page_url") or "")[:600] or None)
    return None, None


def as_published(p, pd, msg, work):
    """The preprint p as the paper it was published as (pd, its registry record msg, its OpenAlex record work), or None
    when that can't be done faithfully: its title must be the preprint's (most words shared), its own abstract and a free
    copy of it must be there to read, so every quote is then the published paper's own sentence."""
    title = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", " ".join(str(x) for x in ((msg or {}).get("title") or [])[:1]))).strip().rstrip(".")
    if not title or not same_title(title, p.get("title")):
        return None
    abstract = published_abstract(msg, work)
    pdf, landing = published_copy(work)
    if not abstract or not pdf:
        return None
    q = dict(p)
    q.update(title=title[:400], doi=pd, url="https://doi.org/" + pd, pdf=pdf, landing=landing, abstract=abstract[:2000],
             abstract_full=abstract[:9000], abstract_lent=False, preprint=False, preprint_doi=p.get("doi"),
             kind="review" if str((work or {}).get("type") or "") == "review" else "article")
    for k in ("pmcid", "bg", "sents", "biblio"):
        q.pop(k, None)
    return q


def registry_check(papers, get, this_year, workers=8, budget=8.0, clock=None):
    """The papers that may be cited, each with its reference data as the DOI's registry holds it, and what was found.
    get(url, params) -> (status, json or None): the server's own fetch (a made-up one in the tests); it must give up after
    a few seconds by itself. Never raises: when the registries can't be reached, the papers keep the index's data and
    stats["unreached"] says so."""
    import time as _time
    clock = clock or _time.time
    end = clock() + budget
    stats = {"checked": 0, "corrected": 0, "retracted": 0, "concern": 0, "wrong_doi": 0, "unknown_doi": 0, "unidentified": 0,
             "future": 0, "unreached": False, "erratum": 0}
    papers = [p for p in (papers or []) if isinstance(p, dict)]
    dois = sorted({p["doi"] for p in papers if p.get("doi") and p.get("kind") != "mine"})
    retracted = set()
    # 1. retractions, whichever index found the paper (one request for every DOI)
    if dois:
        try:
            status, js = get("https://api.openalex.org/works", {"filter": "doi:" + "|".join("https://doi.org/" + d for d in dois[:50]),
                                                                  "select": "doi,is_retracted", "per-page": 50})
            if status == 200 and isinstance(js, dict):
                for w in js.get("results") or []:
                    d = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", str((w or {}).get("doi") or ""), flags=re.I).lower()
                    if w.get("is_retracted") is True and d:
                        retracted.add(d)
            else:
                stats["unreached"] = True
        except Exception:
            stats["unreached"] = True

    # 2. each DOI's own record (Crossref; a DOI another agency gave is looked up at doi.org)
    def one(doi):
        if clock() > end:
            return doi, "late", None
        try:
            status, js = get("https://api.crossref.org/works/" + doi, None)
        except Exception:
            return doi, "error", None
        if status == 200 and isinstance(js, dict) and isinstance(js.get("message"), dict):
            return doi, "crossref", js["message"]
        if status == 404:
            try:
                s2, h = get("https://doi.org/api/handles/" + doi, None)
            except Exception:
                return doi, "error", None
            if s2 == 200 and isinstance(h, dict) and h.get("responseCode") == 1:
                return doi, "other", None
            if s2 == 404 or (isinstance(h, dict) and h.get("responseCode") == 100):
                return doi, "none", None
            return doi, "error", None
        return doi, "error", None

    records = {}
    if dois:
        pool = _futures.ThreadPoolExecutor(max_workers=max(1, workers))
        try:
            jobs = [pool.submit(one, d) for d in dois]
            done, late = _futures.wait(jobs, timeout=max(0.5, end - clock() + 0.5))
            for j in done:
                try:
                    d, kind, msg = j.result()
                except Exception:
                    continue
                records[d] = (kind, msg)
            if late or any(k in ("late", "error") for k, _ in records.values()):
                stats["unreached"] = stats["unreached"] or bool(late) or any(k == "error" for k, _ in records.values())
        finally:
            pool.shutdown(wait=False, cancel_futures=True)

    # 3. 6.6, round 3: a preprint its registry says was published since: the published paper is asked for (its record,
    #    whether it was retracted, its abstract and a free copy of it), and cited instead when it can be read
    stats.update(published=0, same_work=0)
    pub_of = {d: _cr_preprint_of(msg) for d, (kind, msg) in records.items() if kind == "crossref"}
    pub_of = {d: x for d, x in pub_of.items() if x and x != d}
    pub = {}
    wanted = sorted({x for x in pub_of.values() if x not in dois})
    if wanted and clock() < end:
        works = {}
        try:
            status, js = get("https://api.openalex.org/works", {"filter": "doi:" + "|".join("https://doi.org/" + x for x in wanted[:50]),
                                                                  "select": "doi,is_retracted,abstract_inverted_index,best_oa_location,type", "per-page": 50})
            if status == 200 and isinstance(js, dict):
                for w in js.get("results") or []:
                    if isinstance(w, dict) and _doi(w.get("doi")):
                        works[_doi(w.get("doi"))] = w
        except Exception:
            pass
        pool = _futures.ThreadPoolExecutor(max_workers=max(1, workers))
        try:
            jobs = [pool.submit(one, x) for x in wanted]
            done, _late = _futures.wait(jobs, timeout=max(0.5, end - clock() + 0.5))
            for j in done:
                try:
                    x, kind, msg = j.result()
                except Exception:
                    continue
                if kind == "crossref":
                    pub[x] = (msg, works.get(x))
        finally:
            pool.shutdown(wait=False, cancel_futures=True)

    kept = []
    for p in papers:
        d = p.get("doi")
        if p.get("kind") == "mine":
            kept.append(p)
            continue
        if d and d in retracted:
            stats["retracted"] += 1
            continue
        kind, msg = records.get(d, (None, None)) if d else (None, None)
        pd = pub_of.get(d) if d else None
        if pd and pd in dois:
            stats["same_work"] += 1             # its published version is among the papers already: that one is cited
            continue
        if pd and pd in pub:
            pmsg, work = pub[pd]
            gone = _cr_withdrawn(pmsg) or ("retracted" if (work or {}).get("is_retracted") is True else "")
            if gone:
                stats[gone] += 1                # the published paper was retracted (or is under concern): so is the work
                continue
            q = as_published(p, pd, pmsg, work)
            if q is not None:
                p, d, kind, msg = q, pd, "crossref", pmsg
                stats["published"] += 1
        if kind == "none":
            stats["unknown_doi"] += 1          # no such DOI anywhere: the record that gave it can't be trusted
            continue
        if kind == "crossref":
            if any(str((u or {}).get("type") or "").lower() in _CR_RETRACTED for u in (msg.get("updated-by") or [])) or \
                    str(msg.get("subtype") or "").lower() in _CR_RETRACTED:
                stats["retracted"] += 1
                continue
            if any(str((u or {}).get("type") or "").lower() in _CR_CONCERN for u in (msg.get("updated-by") or [])):
                stats["concern"] += 1           # its journal has said the paper's results may not be reliable
                continue
            cr_title = " ".join(str(x) for x in (msg.get("title") or [])[:1])
            if cr_title and not same_title(cr_title, p.get("title")):
                stats["wrong_doi"] += 1         # the DOI is another paper's: which one the index meant can't be known
                continue
            stats["checked"] += 1
            cy = _cr_corrected(msg)                 # 6.6, round 94: a published correction to it
            if cy is not None:
                p["corrected"] = cy or True
                stats["erratum"] += 1
            changed = False
            y = _cr_year(msg)
            if y and y != p.get("year"):
                p["year"], changed = y, True
            auth = _cr_authors(msg)
            if auth and [a.lower() for a in auth] != [str(a).lower() for a in (p.get("authors") or [])]:
                p["authors_registry"], changed = auth, True
            venue = " ".join(str(x) for x in (msg.get("container-title") or [])[:1]).strip()
            if venue and str(msg.get("type") or "") in ("journal-article", "proceedings-article", "book-chapter") and venue != p.get("venue"):
                p["venue"], changed = venue[:200], True
            # 6.6, round 67: the journal's abbreviation, as its publisher registered it (NLM's Vancouver names a journal so)
            short = journal_abbrev(msg)
            if short:
                p["journal_abbrev"] = short
            b = dict(p.get("biblio") or {})
            for k, v in (("volume", str(msg.get("volume") or "").strip()), ("issue", str(msg.get("issue") or "").strip())):
                if v and b.get(k) != v:
                    b[k], changed = v[:20], True
            pg = _cr_pages(msg)
            if pg and (pg.get("first_page") != b.get("first_page") or pg.get("last_page", "") != (b.get("last_page") or "")):
                b.pop("last_page", None)
                b.update(pg)
                changed = True
            p["biblio"] = b
            # 6.6, round 61: a chapter is a part of a book, written as one (its book's title is the venue above), with its
            # book's editors and publisher; a book with its publisher
            ctype = str(msg.get("type") or "").lower()
            if ctype in _CR_CHAPTER:
                p["chapter"] = True
                eds = _cr_authors(msg, "editor")
                if eds:
                    p["editors"] = eds[:12]
            if ctype in _CR_BOOKISH:
                pub = re.sub(r"\s+", " ", str(msg.get("publisher") or "")).strip()
                if pub:
                    p["publisher"] = pub[:160]
            p["registry"] = "crossref"
            stats["corrected"] += 1 if changed else 0
        elif kind == "other":
            p["registry"] = "doi"
            stats["checked"] += 1
        if not d and not (str(p.get("venue") or "").strip() or str(p.get("institution") or "").strip()):
            stats["unidentified"] += 1          # neither a DOI nor a journal, publisher or institution: who published it?
            continue
        if p.get("year") and int(p["year"]) > this_year:
            stats["future"] += 1                # a year after this one is wrong, and nothing gave the right one
            continue
        kept.append(p)
    return kept, stats


def checked_words(stats, n_papers):
    """What the person is told about the papers' check, in plain words ("" when there is nothing to say)."""
    if not stats:
        return ""
    out = []
    gone = stats.get("retracted", 0)
    if gone:
        out.append(f"{gone} retracted paper{'' if gone == 1 else 's'} left out")
    worry = stats.get("concern", 0)
    if worry:
        out.append(f"{worry} paper{'' if worry == 1 else 's'} under an expression of concern left out")
    bad = stats.get("wrong_doi", 0) + stats.get("unknown_doi", 0) + stats.get("unidentified", 0) + stats.get("future", 0)
    if bad:
        out.append(f"{bad} paper{'' if bad == 1 else 's'} whose details couldn't be confirmed left out")
    if stats.get("unreached"):
        out.append("the publishers' records couldn't all be reached this time")
    return "; ".join(out)
