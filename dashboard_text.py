# -*- coding: utf-8 -*-
"""
Plain-language layer for the Air Quality dashboard.

Every chart gets: a question-style title, a one-line description in everyday
words, and (for the harder ones) a "how to read it" note. The Roman Urdu text
(`ur`) is not shipped in the dashboard: make_guide.py uses it to write the
walk-through document.
"""
import re

TABS = {  # short labels so all ten tabs fit on a laptop screen
    "overview": "📊 Overview", "profile": "👨‍👩‍👧 Respondents", "child": "🎒 Child & Health",
    "classes": "🏫 Classrooms", "aware": "🌫️ Awareness", "wtp": "💰 Willing to Pay", "effects": "🧪 Video Effects", "tracker": "📋 Tracker",
}
SHORT = {  # one line per section for the menu
    "overview": "Progress, study groups and headline findings",
    "classes": "Which classes reach the 30% rule and get a purifier",
    "profile": "Who was interviewed",
    "child": "School, commute and child health",
    "aware": "What parents know and believe about air pollution",
    "wtp": "How much parents would pay for a classroom purifier",
    "effects": "Did the videos change what parents say?",
    "tracker": "The status of every household",
}
PANEL_TITLE = {"effects": "Did the Videos Make a Difference?"}

DK = re.compile(r"(?i)don'?t know|refused|not applicable|not sure|^unsure")

# --------------------------------------------------------------------------- hard-chart explainers (shown in the ? panel)
HELP = {
    "group": {
        "read": "Each bar is the <strong>average</strong> amount that group of parents said they would pay. Compare the bars: a longer bar means that group is willing to pay more on average. "
                "The number after each bar is how many parents are in the group.",
        "care": "Groups with only a few parents can swing a lot. Look at n before drawing a conclusion.",
    },
    "likert": {
        "read": "Each row is one statement. The bar is split into the share of parents giving each answer, from left to right. "
                "Add up the two leftmost colours to see the share who agree (or, for the score chart, who expect the biggest gains).",
        "care": "Neutral (grey) answers sit in the middle: a long grey block means many parents had no strong view.",
    },
    "bid": {
        "read": "Each parent named the most they would contribute towards an air purifier for their child's classroom. A random price is then drawn: "
                "if the bid is at least as high they pay the random price, otherwise nothing. That makes the honest bid the best bid. Each bar counts parents whose bid fell in that band.",
        "care": "The tall bar at Rs 2,000 is the maximum allowed: some parents would have paid more.",
    },
    "cum": {
        "read": "The green line is the number of completed interviews so far. The dashed line continues it at the average pace to date. The red dashed line is the target. "
                "Where the dashed line reaches the red one is the expected finish date.",
        "care": "It assumes the team keeps today's average pace and skips Sundays.",
    },
    "others": {
        "read": "Each parent guessed how the other parents in their class would behave. Bars show the average guess, as a share of the class, for each price band. "
                "“Not surveyed” is the share the parent thought we would not reach.",
        "care": "Guesses are opinions, not facts.",
    },
    "classrule": {
        "read": "A purifier is installed in a class only if at least 30% of that class's parents end up contributing. Each parent's contribution is decided by a random price drawn against their own bid: "
                "if the random price is at or below the bid, the parent contributes it. For a class of 30 parents, 9 contributors are enough. Parents who have not been interviewed (or cannot be reached) count as not contributing.",
        "care": "Because prices are drawn at random, two classes with the same average bid can end up on different sides of the line. A class is only final once every parent has been interviewed.",
    },
    "chance": {
        "read": "“Chance of reaching 30%” looks at the parents still to be interviewed. It assumes each of them contributes with the same probability as the class average so far, "
                "and adds up how likely it is that enough of them will contribute to bring the class to 30%. Above 50% is “On track”; below is “At risk”.",
        "care": "It is a projection from current bids, not a promise. It sharpens as more parents in the class are interviewed.",
    },
}

GENERIC = {  # default explanation by chart type (used when a chart has no specific note)
    "donut": "Each slice is one answer. The bigger the slice, the more parents chose it. The slices add up to 100%.",
    "bar": "Each bar is one group or answer. The taller the bar, the more parents fall in it. The number on top is the percentage of parents.",
    "hbar": "Each bar is one answer. The longer the bar, the more parents chose it. The number at the end is the percentage of parents.",
    "likert": "Each row is one statement, split into the share of parents giving each answer.",
    "lines": "The line follows the numbers over time or across prices. Hover a point to see the exact value.",
    "gauge": "The green part shows how much of the target is already done.",
    "table": "Click a column heading to sort. Use the buttons above the table to filter.",
    "funnel": "The first bar is everything we started with. Each next bar is what remains after that step.",
    "blocks": "Each card is one part of the sample. The bar shows how full it is.",
    "calendar": "Each box is a day. Darker green means more completed interviews that day.",
}

# --------------------------------------------------------------------------- per-chart text
# id: (title, one-line description, Roman Urdu explanation, help-key or None)
def card_text(target):
    T = target
    return {
        # ---- overview
        "ovDaily": ('Interviews completed each day', 'Number of completed interviews on each field day.',
                    'Har din kitne interviews mukammal hue. Line upar ja rahi ho to team ki raftaar barh rahi hai.', None),
        "ovProg": ("How close are we to the target?", f"Share of the {T}-household sample already completed.",
                   "Gol chart batata hai target ka kitna hissa poora ho chuka. Hara hissa mukammal, grey hissa baaqi. Beech mein % aur ginti likhi hai.", None),
        "ovArm": ("Interviews by study group", "How many interviews are done in Control (no video), Video 1 and Video 2.",
                  "Teeno groups mein kitne interviews hue: Control (video nahi), Video 1, Video 2. Teeno ka lagbhag barabar bharna zaroori hai taake muqabla fair rahe.", None),
        "ovDisp": ('What happened at each visit', 'Outcome of every visit, with the number and the share of all visits.',
                    'Har visit ka nateeja: mukammal, inkaar, ghar band wagera. Har bar ke saath ginti aur % likha hai. Agar inkaar zyada ho to parents se baat karne ka tareeqa behtar karna parta hai.', None),
        "ovFunnel": ("From first visit to a usable interview", "How many visits are lost at each step.",
                     "Pehli bar saare visits, phir jo mile aur razi hue, phir mukammal, phir analysis mein gine gaye. Har agli bar choti hoti hai kyunke kuch cases beech mein nikal jate hain.", None),
        "ovBlocks": ("How is each school progressing?", "Completed interviews against the parents listed in each school's classes.",
                     "Har school ka card: kitne parents ki list thi aur kitno ka interview ho chuka. Bar jitni bhari utna kaam mukammal. Jo school peeche ho wahan zyada visits karwaein.", None),
        # ---- classrooms
        "clStatus": ("Where do the classes stand?", "Number of classes in each status.",
                     "Kitni classes 30% tak pahunch chuki hain (Secured), kitni pahunchne wali hain (On track), kitni khatre mein (At risk), aur kitni ab pahunch hi nahi sakti (Cannot reach). Har class ko alag alag dekha jata hai, poore sample ko nahi.", "classrule"),
        "clDist": ("How many classes are near the target?", "Classes grouped by how much of the contributors they need they already have.",
                   "Har class ko purifier ke liye kuch contributors chahiyein (class ka 30%). Yeh graph batata hai kitni classes ke paas zaroori contributors ka 0-25%, 25-50%, 50-75%, 75-100% aur poora (100%+) hai. Aakhri (hara) group wali classes ko purifier mil jata hai.", None),
        "clBars": ("How close is each class to the number it needs?", "Contributors so far as a share of the number needed (30% of the class). The dashed line is 100%: secured.",
                   "Har class ek bar hai. Bar batati hai ke class ko jitne contributors chahiyein (class ka 30%) un mein se kitne ho chuke. Bar ke aakhir mein likha hai, jaise '9 of 12': 12 chahiyein, 9 ho gaye. Dashed line 100% hai: jo bar us se aagay nikal jaye us class ko purifier mil jata hai. Jinka interview nahi hua unhein contribute nahi gina jata. Rang status batata hai: hara = secured, neela = on track, narangi = at risk, laal = cannot reach.", "classrule"),
        "clSchool": ("How does each school's classes split?", "Each bar is one school; the colours are its classes' status.",
                     "Har school ki classes kin kin status mein hain. Number classes ki ginti hai. Is se pata chalta hai kaun se school ki classes zyada tar purifier hasil kar rahi hain.", None),
        "clArm": ("Do the videos raise the share who contribute?", "Share whose random price cleared their bid, by study group.",
                  "Control, Video 1 aur Video 2 mein kitne % parents ka random price unke bid se kam nikla (yani woh contribute karenge). Video ka asar dekhne ka seedha tareeqa.", None),
        "clTable": ("Every class, one line each", "Click a column to sort, or filter by school. “Chance” is the probability of reaching 30% once the parents still to interview are counted.",
                    "Har class ki poori tafseel: class size, kitne parents ka interview hua, kitne contribute kar rahe hain, kitne chahiyein (30%), aur baaqi parents ke baad 30% tak pahunchne ka imkaan (Chance). Status aakhri column mein hai.", "chance"),
        # ---- field ops
        "opsCum": ('Are we on pace to finish?', "Interviews done so far, where today's pace takes us (dashed), and the target (red).",
                    'Hari line ab tak ke mukammal interviews hain. Dashed line maujooda raftaar par agla andaza hai aur laal line target. Jahan dashed line laal line ko chhoti hai wahi mutawaqqa khatam hone ki tareekh hai.', "cum"),
        # ---- profile
        "prArea": ("Where do the households live?", "Urban, peri-urban or rural.", "Ghar shehar mein, shehar ke kinare (peri-urban) ya dehaat mein hain. Har slice ka hissa % mein.", None),
        "prGender": ("Who answered the survey?", "Gender of the parent or caregiver interviewed.", "Jawab dene wale walid ya walida ki jins. Aam tor par mothers ziyada hoti hain.", None),
        "prDec": ("Who decides about the child's schooling?", "The person who usually makes schooling decisions.", "Bachay ki taleem ka faisla kaun karta hai: maa, baap ya dono.", None),
        "prAge": ("How old are the parents?", "Age of the person interviewed, in groups.", "Jawab dene walon ki umar ke groups. Sabse lambi bar sabse aam umar ka group hai.", None),
        "prInc": ("How much does the household earn each month?", "Monthly household income in rupees, in groups.", "Ghar ki mahana aamdani (rupay) ke groups. Yeh samajhne ke liye ke kis tabqe se parents hain.", None),
        "prSize": ("How many people live in the home?", "Household size, in groups.", "Ghar mein kitne log rehte hain, groups mein.", None),
        "prAssets": ("What does the household own?", "Share of households owning each item.",
                     "Kitne % gharon mein AC, mobile, internet aur air purifier hai. Air purifier bohat kam gharon mein hai, yani parents ek naye product ki qeemat laga rahe hain.", None),
        "prWin": ("Were windows open?", "What the enumerator saw when entering the home.", "Enumerator ne ghar mein dakhil hote waqt khirkiyan khuli dekhi ya band.", None),
        "prIll": ("Anyone with a breathing illness at home?", "Respiratory illness in the household.", "Ghar mein kisi ko saans ki bimari hai ya nahi.", None),
        "prIll2": ("Is it a child or an adult?", "Who is affected, where illness is reported.", "Agar bimari hai to bachay ko hai ya bade ko.", None),
        # ---- child
        "chGrade": ('Which class is the child in?', 'Class of the child in the study classroom.',
                    'Bachay ki class.', None),
        "chExam": ("How did the child do in the last exam?", "Most recent exam result reported by the parent, in percent.", "Pichlay imtihan mein bachay ke numbers (percent). Sirf wahi jawab gine gaye jin mein percentage thi.", None),
        "chTime": ("How long is the journey to school?", "One-way travel time in minutes.", "Ghar se school tak ka ek taraf ka safar kitne minute ka hai.", None),
        "chTravel": ("How do children get to school?", "The main way the child travels.", "Bachay school kaise jate hain: paidal, rickshaw/van, gaari, school transport wagera.", None),
        "chCond": ("Which long-term conditions did children have?", "Share of children with each condition in the past two months (a child can have more than one).",
                   "Pichlay 2 mahinon mein bachay ko kaun si lambi bimari rahi (asthma, saans, dil wagera). Ek bachay ko ek se zyada bhi ho sakti hai is liye total 100% se zyada ho sakta hai.", None),
        "chDays": ("How many school days were missed?", "Days the child could not go to school because of illness, past two months.", "Bimari ki wajah se bachay ne kitne din school miss kiya.", None),
        "chAir": ("Do parents link illness to air pollution?", "Yes, no or don't know.", "Kya parents bachay ki bimari ko hawa ki aloodgi se jorte hain.", None),
        "chTut": ("Do children go to tuition?", "Tuition or academy after school.", "Kya bachay school ke baad tuition ya academy jata hai.", None),
        # ---- awareness
        "awAqi": ("What do parents think the AQI means?", "Parents could pick more than one answer.", "Parents ke nazdeek AQI kya hai. 'Levels of air pollution' sahi jawab hai; baaqi ghalat fehmi hain.", None),
        "awRate": ("How polluted do parents think the air is?", "0 = not polluted at all, 10 = extremely polluted.", "Parents ke mutabiq bahar ki hawa kitni gandi hai (0 se 10). Bars jitni daayen taraf hon utna ziyada gandi samajhte hain.", None),
        "awCheck": ("How often do parents check air quality?", "Checking a website or app for pollution levels.", "Parents kitni baar hawa ka AQI check karte hain: rozana, hafta, kabhi nahi.", None),
        "awSource": ("Where do parents get air-quality news?", "Main source of information.", "Parents ko hawa ki khabar kahan se milti hai: TV, social media, apps, dost.", None),
        "awIndoor": ("Is the air at home cleaner than outside?", "How polluted parents think the home is compared with outdoors.", "Parents ke khayal mein ghar ke andar ki hawa bahar se kitni gandi ya saaf hai.", None),
        "awRisk": ("How much harm do parents think pollution does?", "Believed increase in the child's risk of breathing illness.", "Parents ke mutabiq aloodgi se bachay ki saans ki bimari ka khatra kitna barhta hai (koi asar nahi se lekar dugna tak).", None),
        "awExam": ("Would a purifier or an air conditioner raise exam marks?", "How much higher parents think their child would score with each in the classroom.",
                   "Parents ke khayal mein classroom mein AC ya air purifier hone se numbers kitne barh jayenge. Jitna gehra rang utna ziyada faida. Purifier ka rang ziyada gehra ho to parents usay AC se behtar samajhte hain.", "likert"),
        "awKnow": ("Ways parents know to reduce the harm", "Answers were not read out; parents named what they know.", "Parents kaun se tareeqe jaante hain jin se aloodgi ka nuqsan kam ho (mask, khirkiyan band, purifier wagera).", None),
        "awDo": ("What parents actually do to protect the child", "Answers were not read out; parents named what they do.", "Parents bachay ko bachane ke liye asal mein kya karte hain.", None),
        "awSeason": ("Would the same child score lower in November?", "Same preparation, same child: an April test versus a November test.", "Wahi bacha, wahi tayyari, bas imtihan April ke bajaye November mein ho (smog ka mausam). Kya numbers kam ayenge?", None),
        "awSchool": ("Does the school do anything about air quality?", "Fans, air conditioners, an open-window policy or similar.", "Kya school mein hawa behtar karne ke liye kuch hai.", None),
        "awRank": ("How important is school air quality?", "Compared with better teachers, facilities or books.", "School ki hawa ki ahmiyat doosri cheezon (ustaad, facilities, kitaabein) ke muqablay mein kitni hai. 1 = bohat kam, 5 = bohat ziyada.", None),
        # ---- willingness to pay
        "wtDist": ("How much would parents pay?", "The most each parent would contribute, after any change, in rupees.",
                   "Har parent ne kitne rupay dene ka kaha, groups mein. Lambi bar = ziyada parents. Rs 2,000 par bar isliye hoti hai ke woh had (maximum) thi.", "bid"),
        "wtDemand": ('How many parents would pay at least this much?', 'Share of parents whose bid is at or above each amount.',
                    'Har bar ek raqam hai (Rs 100, 250, 500...). Bar batati hai kitne % parents ne itne ya us se zyada ka bid diya. Raqam barhti hai to bar chhoti hoti jati hai.', None),
        "wtGroup": ('Average amount parents would pay, by group', 'Average final bid for each group; not affected by the filter above.',
                    'Alag alag groups (area, jins, aamdani wagera) ka average bid. Bar jitni lambi, us group ke parents utne zyada dene ko tayyar. Bar ke saath n likha hai: group mein kitne parents hain.', "group"),
        "wtOthers": ("What parents expect other parents to pay", "Average guess of how the other parents in the class will behave.", "Har parent ne andaza lagaya ke class ke doosre parents kitna denge. Bars unka average andaza hain (class ka hissa %).", "others"),
        "wtChange": ("Did parents change their bid?", "After seeing the result of the random draw.", "Random number dekhne ke baad kitne parents ne apna bid badla.", None),
        "wtCertain": ("How sure are parents of their bid?", "Certainty about the amount chosen.", "Parents apne bid par kitne yaqeen mein hain.", None),
        "wtWhy": ("Why not pay more?", "Main reason given by parents who bid below the maximum.", "Jinhon ne Rs 2,000 se kam bid kiya unhon ne ziyada kyun nahi diya.", None),
        "wtPractice": ("Practice round", "A Rs 0–20 practice before the real question.", "Asli sawal se pehle Rs 0 se 20 ki mashq (practice). Yeh sirf samajhne ke liye thi.", None),
        "wtSel": ("Understanding check 1", "If selected and the random price is above the bid, is the Rs 1,000 bonus still received?", "Samajh ki jaanch: kya parents ne bonus wala rule sahi samjha.", None),
        "wtNot": ("Understanding check 2", "If not selected, which bid is used?", "Samajh ki jaanch: agar parent muntakhib na ho to kaun sa bid use hota hai.", None),
        # ---- effects
        "teArmCi": ('Average amount parents would pay, by study group', 'Average final bid in Control, Video 1 and Video 2.',
                    'Control (video nahi), Video 1 aur Video 2 ka average bid. Jo bar lambi hai us group ke parents ziyada dene ko tayyar hain.', "group"),
        "teEff": ('Did the videos change what parents say?', 'Average for each group. The highest number in each row is highlighted.',
                    'Har row ek sawal hai aur teen columns mein Control, Video 1, Video 2 ka average ya %. Jis group ka number sabse zyada ho woh hara highlight hota hai. Is se seedha nazar aata hai kis video ka asar zyada hai.', None),
        "teBal": ('Are the three groups alike?', 'Background of the parents in each group. Similar numbers mean a fair comparison.',
                    'Muqabla tabhi fair hai jab teeno groups shuru mein ek jaise hon (umar, aamdani, area). Numbers qareeb hon to theek hai.', None),
        "teOrder": ('Does asking about other parents first change the bid?', 'Average bid when parents were asked about other parents before or after their own bid.',
                    'Kuch parents se doosron ke bare mein pehle poochha gaya aur kuch se bid ke baad. Yeh dekhne ke liye ke sawal ka silsila jawab badalta hai ya nahi.', "group"),
        # ---- fairness
        "faAgree": ('Who should pay for classroom air purifiers?', 'How much parents agree with each statement.',
                    'Teen bayanaat par parents kitne muttafiq hain. Gehra neela = zor se agree, laal = disagree.', "likert"),
        # ---- time
        "tpPaid": ("How much did parents actually pay?", "Under the random-price rule; zero if the price drawn was above the bid.", "Random price rule ke mutabiq parents ne asal mein kitna diya. Zero matlab random price bid se ziyada nikla.", None),
    }


# --------------------------------------------------------------------------- one-line "in plain words" read-out
def auto_read(kind, d, opt, cid=None):
    opt = opt or {}
    try:
        if kind in ("bar", "hbar", "donut"):
            if opt.get("grouped"):
                rows = [r for r in d["rows"] if r.get("m") is not None]
                if len(rows) < 2:
                    return None
                hi = max(rows, key=lambda r: r["m"]); lo = min(rows, key=lambda r: r["m"])
                nm = lambda r: f"{r['g']}: {r['l']}" if r["g"] not in ("Study group", "Control", "Video 1", "Video 2") else (r["l"] if r["g"] == "Study group" else f"{r['g']}, {r['l'].lower()}")
                return f"Highest average: <strong>{nm(hi)}</strong> (Rs {hi['m']:,.0f}). Lowest: <strong>{nm(lo)}</strong> (Rs {lo['m']:,.0f})."
            its = d["items"]; fmt = opt.get("fmt", "pct")
            key = {"n": "n", "num1": "v", "v": "v"}.get(fmt, "p")
            cand = [i for i in its if not DK.search(i["l"]) and i.get(key) is not None]
            if not cand:
                return None
            top = max(cand, key=lambda i: i[key])
            if opt.get("tpl"):
                return opt["tpl"].format(l=top["l"], p=top.get("p"), n=top.get("n"), v=top.get("v"))
            if fmt == "n":
                return f"Largest: <strong>{top['l']}</strong> with {top['n']:,}."
            if fmt == "num1":
                return f"Largest expected share: <strong>{top['l']}</strong> ({top['v']:.0f}% of the class on average)."
            if d.get("hist"):
                return f"Most parents fall in <strong>{top['l']}</strong>: {top['p']:.0f}% (n = {top['n']})."
            return f"Most common answer: <strong>{top['l']}</strong>, {top['p']:.0f}% of parents (n = {top['n']})."
        if kind == "likert":
            cats = d["cats"]; seq = opt.get("pal") == "seq"
            idx = [len(cats) - 2, len(cats) - 1] if seq else [0, 1]
            lead = opt.get("lead") or ("Expect a large gain (top two answers): " if seq else "Agree or strongly agree: ")
            return lead + " · ".join(f"{r['l'].split(':')[0][:48]} <strong>{sum(r['p'][i] for i in idx):.0f}%</strong>" for r in d["rows"])
        if kind == "gauge":
            left = max(d["target"] - d["done"], 0)
            return f"<strong>{d['done']}</strong> of <strong>{d['target']}</strong> households completed; <strong>{left}</strong> to go."
        if kind == "lines":
            if opt.get("ytitle"):
                ss = [s for s in d["series"] if not s.get("dash")]
                if len(ss) == 1:
                    y = ss[0]["y"]
                    ok = [i for i, v in enumerate(y) if v is not None and v >= 30]
                    tail = f" The 30% line is still met up to about <strong>Rs {ok[-1] * 100:,}</strong>." if ok else " Fewer than 30% would pay even the lowest price."
                    return f"<strong>{y[5]:.0f}%</strong> of parents would pay at least Rs 500 and <strong>{y[10]:.0f}%</strong> at least Rs 1,000." + tail
                return "Share still willing to pay Rs 1,000 or more: " + " · ".join(f"{s['name'].split(' (')[0]} <strong>{s['y'][10]:.0f}%</strong>" for s in ss if s["y"][10] is not None)
            s0 = d["series"][0]
            if s0["name"] == "Completed":
                ys = [(v, x) for v, x in zip(s0["y"], d["x"]) if v is not None]
                if ys:
                    best = max(ys)
                    return f"Busiest day: <strong>{best[1]}</strong> with {best[0]} completed. Average {sum(v for v, _ in ys) / len(ys):.1f} per field day."
    except Exception:
        return None
    return None
