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
    "overview": "📊 Overview", "ops": "🛠️ Field Ops", "profile": "👨‍👩‍👧 Respondents", "child": "🎒 Child & Health",
    "aware": "🌫️ Awareness", "wtp": "💰 Willing to Pay", "effects": "🧪 Video Effects", "fair": "⚖️ Fairness",
    "time": "⏳ Patience", "tracker": "📋 Tracker",
}
SHORT = {  # one line per section for the menu
    "overview": "Progress, study groups and headline findings",
    "ops": "Pace, team performance and data checks",
    "profile": "Who was interviewed",
    "child": "School, commute and child health",
    "aware": "What parents know and believe about air pollution",
    "wtp": "How much parents would pay for a classroom purifier",
    "effects": "Did the videos change what parents say?",
    "fair": "Who should pay? Free-riding and priorities",
    "time": "Patience and the mobile top-up",
    "tracker": "The status of every household",
}
PANEL_TITLE = {"effects": "Did the Videos Make a Difference?", "time": "Patience & Mobile Top-up"}

DK = re.compile(r"(?i)don'?t know|refused|not applicable|not sure|^unsure")

# --------------------------------------------------------------------------- hard-chart explainers (shown in the ? panel)
HELP = {
    "ci": {
        "read": "Each bar is the <strong>average</strong> amount that group of parents said they would pay. The thin line with two end-caps is the "
                "<strong>likely range</strong> (a 95% confidence interval). It says: if we repeated this survey many times, the true average would fall "
                "inside this range about 95 times out of 100. A <strong>short line</strong> means we are fairly sure of the average. "
                "A <strong>long line</strong> means few parents answered, so the average could easily be a bit higher or lower.",
        "care": "If the thin lines of two groups overlap a lot, the groups are <strong>not clearly different</strong>, even if one bar looks longer. "
                "Trust a difference only when the lines barely overlap.",
    },
    "demand": {
        "read": "Move along the bottom from cheap to expensive. The line shows what share of parents would still pay <strong>that much or more</strong>. "
                "It always starts at 100% (everybody will pay Rs 0) and falls as the price rises. The dashed line marks 30%: "
                "a classroom only gets its purifier if 30% of parents contribute, so the point where the red line crosses the dashed line is the highest price that still works.",
        "care": "A line that sits higher means parents value the purifier more. Two lines close together mean no real difference.",
    },
    "effects": {
        "read": "Each row compares two groups of parents on one question. <strong>Difference</strong> = treatment group minus comparison group. "
                "<strong>Likely range</strong> is the 95% confidence interval: the gap is probably somewhere inside it. "
                "The <strong>p-value</strong> is the chance of seeing a gap this big by luck if the video changed nothing: small (below 0.05) means it is unlikely to be luck. "
                "The last column puts that into words.",
        "care": "If the likely range includes zero, we cannot say the video changed anything. With about 33 households per group the ranges are wide; "
                "they narrow as fieldwork completes. Many rows are tested at once, so one or two stray “clear differences” can still happen by chance.",
    },
    "balance": {
        "read": "Before comparing groups we check that the three groups of parents looked alike to begin with (age, income, area and so on). "
                "Because the groups were drawn by lottery they should. A p-value below 0.10 is flagged with ⚑ so the team can double-check.",
        "care": "Some flags are expected by pure chance (about 1 in 10 rows). A flag is a prompt to look, not proof of a problem.",
    },
    "discount": {
        "read": "Parents chose between Rs 2,000 today and a bigger amount in one year. The amount at which they switch shows how much extra they want for waiting. "
                "Example: switching at Rs 4,000 means asking for a 100% yearly return. Higher bars on the right mean more impatient parents.",
        "care": "Very high rates are common when money is tight; they mean people strongly prefer cash now.",
    },
    "likert": {
        "read": "Each row is one statement. The bar is split into the share of parents giving each answer, from left to right. "
                "Add up the two leftmost colours to see the share who agree (or, for the score chart, who expect the biggest gains).",
        "care": "Neutral (grey) answers sit in the middle: a long grey block means many parents had no strong view.",
    },
    "bid": {
        "read": "Each parent named the most they would contribute towards an air purifier for their child's classroom. A random price is then drawn: "
                "if the bid is at least as high they pay the random price, otherwise nothing. That makes the honest bid the best bid. Each bar counts parents whose bid fell in that band.",
        "care": "The spike at Rs 2,000 is the maximum allowed: some parents would have paid more.",
    },
    "qa": {
        "read": "The dashboard runs automatic checks on every interview. “Verify” means the field team should look at those records again. “Clear” means nothing found.",
        "care": "A flag is a prompt to look, never a finding that someone did anything wrong.",
    },
    "cum": {
        "read": "The green line is the number of completed interviews so far. The dashed line continues it at the average pace to date. The red dashed line is the target. "
                "Where the dashed line reaches the red one is the expected finish date.",
        "care": "It assumes the team keeps today's average pace and skips Sundays.",
    },
    "others": {
        "read": "Each parent guessed how the other parents in their class would behave. Bars show the average guess, as a share of the class, for each price band. "
                "“Not surveyed” is the share the parent thought we would not reach.",
        "care": "Guesses are opinions, not facts. Compare them with the real bids on the demand curve.",
    },
}

GENERIC = {  # default explanation by chart type (used when a chart has no specific note)
    "donut": "Each slice is one answer. The bigger the slice, the more parents chose it. The slices add up to 100%.",
    "bar": "Each bar is one group or answer. The taller the bar, the more parents fall in it. The number on top is the percentage of parents.",
    "hbar": "Each bar is one answer. The longer the bar, the more parents chose it. The number at the end is the percentage of parents.",
    "likert": "Each row is one statement, split into the share of parents giving each answer.",
    "lines": "The line follows the numbers over time or across prices. Hover a point to see the exact value.",
    "gauge": "The green part shows how much of the target is already done.",
    "table": "Click a column heading to sort. Use the buttons above the table to switch the comparison.",
    "funnel": "The first bar is everything we started with. Each next bar is what remains after that step.",
    "blocks": "Each card is one part of the sample. The bar shows how full it is.",
    "calendar": "Each box is a day. Darker green means more completed interviews that day.",
    "qa": "A list of automatic checks with the number of records to look at again.",
    "recon": "Shows whether the two exported data files (CSV and Stata) contain the same information.",
}

# --------------------------------------------------------------------------- per-chart text
# id: (title, one-line description, Roman Urdu explanation, help-key or None)
def card_text(target):
    T = target
    return {
        # ---- overview
        "ovDaily": ("Interviews completed each day", "Each point is one field day. The dashed line smooths the ups and downs (3-day average).",
                    "Har din kitne interviews mukammal hue. Laal line rozana ginti hai aur dashed line 3 din ka average hai jo utaar chadhaao ko hamwar karti hai. Line upar ja rahi ho to team ki raftaar barh rahi hai.", None),
        "ovProg": ("How close are we to the target?", f"Share of the {T}-household sample already completed.",
                   "Gol chart batata hai target ka kitna hissa poora ho chuka. Hara hissa mukammal, grey hissa baaqi. Beech mein % aur ginti likhi hai.", None),
        "ovArm": ("Interviews by study group", "How many interviews are done in Control (no video), Video 1 and Video 2.",
                  "Teeno groups mein kitne interviews hue: Control (video nahi), Video 1, Video 2. Teeno ka lagbhag barabar bharna zaroori hai taake muqabla fair rahe.", None),
        "ovDisp": ("What happened at each visit", "Outcome of every visit: completed, refused, house locked and so on.",
                   "Har visit ka nateeja: mukammal, inkaar, ghar band wagera. Agar inkaar ziyada ho to team ko parents se baat karne ka tareeqa behtar karna parta hai.", None),
        "ovFunnel": ("From first visit to a usable interview", "How many visits are lost at each step.",
                     "Pehli bar saare visits, phir jo mile aur razi hue, phir mukammal, phir analysis mein gine gaye. Har agli bar choti hoti hai kyunke kuch cases beech mein nikal jate hain.", None),
        "ovBlocks": ("Are all six parts of the sample filling evenly?", "Progress for each group (Control, Video 1, Video 2) and question order.",
                     "Sample 6 hissoon mein banta hai (3 groups x 2 sawalon ka order). Har card dikhata hai wo hissa kitna bhar chuka. Jo hissa peeche ho us par ziyada visits karwaein.", None),
        # ---- field ops
        "opsCum": ("Are we on pace to finish?", "Interviews done so far, where today's pace takes us (dashed), and the target (red).",
                   "Hari line ab tak ke mukammal interviews hain. Dashed line maujooda raftaar par agla andaza hai aur laal line target. Jahan dashed line laal line ko chhoti hai wahi mutawaqqa khatam hone ki tareekh hai.", "cum"),
        "opsCal": ("Which days were busy?", "Each box is a day; darker green means more completed interviews.",
                   "Calendar ki tarah: har dabba ek din. Jitna gehra hara utne zyada interviews. Grey dabbay wo din jab koi complete nahi hua.", None),
        "opsQa": ("Records the team should double-check", "Automatic checks. A flag means look again, not that anything is wrong.",
                  "Automatic jaanch ki list. 'Verify' ka matlab hai dubara dekh lein, ghalat hona zaroori nahi. 'Clear' matlab koi masla nahi mila.", "qa"),
        "opsEnum": ("How is each field team member doing?", "Visits, completed interviews, usual interview length and bid pattern for each person.",
                    "Har enumerator ki table: kitne visits, kitne mukammal, aam interview kitne minute ka, aur unke interviews mein average bid. 'To verify' column mein wajah likhi hoti hai jis par baat karni chahiye.", None),
        "opsDur": ("How long do interviews take?", "Number of interviews in each length band, in minutes.",
                   "Interview ki lambai. Zyada tar 25 se 60 minute ke beech hone chahiye. 15 minute se chhote interviews shak paida karte hain ke sawal jaldi mein poochay gaye.", None),
        "opsHour": ("What time of day do interviews start?", "Number of interviews beginning in each hour of the day (24-hour clock).",
                    "Interviews din ke kis ghante mein shuru hue. Subah 7 se pehle ya raat 9 ke baad wale interviews check karne chahiye.", None),
        "opsEnum2": ("Completed interviews per person", "Each bar is one field team member.",
                     "Har enumerator ke mukammal interviews. Lambi bar matlab ziyada kaam.", None),
        "opsRecon": ("Do the CSV and Stata files agree?", "Compares the two exported data files before they are combined.",
                     "SurveyCTO ki CSV aur Stata (.dta) file ka milaap. 'Cells that differ' zero ho to dono bilkul same hain; warna dashboard cleaned .dta ki value use karta hai.", None),
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
        "chGrade": ("Which class is the child in?", "Class of the child in the study classroom.", "Bachay ki class (6 se 10 tak zyada tar).", None),
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
        "wtDemand": ("How many parents would pay each price?", "Share of parents who would pay at least each price. The dashed line is the 30% a classroom needs.",
                     "Neeche qeemat hai, upar wo % parents jo itni ya us se ziyada qeemat de denge. Qeemat barhti hai to % girta hai. Dashed line 30% hai: classroom ko purifier tab milta hai jab 30% parents contribute karein. Laal line dashed line ke upar rahe to us qeemat par purifier lag sakta hai.", "demand"),
        "wtGroup": ("Average amount parents would pay, by group", "Bar = average bid. Thin line = the likely range. Not affected by the filter above.",
                    "Alag alag groups (area, jins, aamdani wagera) ka average bid. Bar average hai. Patli line 'mumkin range' (confidence interval) hai. Chhoti line = hum average par yaqeen kar sakte hain; lambi line = jawab kam hain is liye andaza kamzor hai. Do groups ki lines ziyada milti hon to unmein wazeh farq nahi.", "ci"),
        "wtOthers": ("What parents expect other parents to pay", "Average guess of how the other parents in the class will behave.", "Har parent ne andaza lagaya ke class ke doosre parents kitna denge. Bars unka average andaza hain (class ka hissa %).", "others"),
        "wtChange": ("Did parents change their bid?", "After seeing the result of the random draw.", "Random number dekhne ke baad kitne parents ne apna bid badla.", None),
        "wtCertain": ("How sure are parents of their bid?", "Certainty about the amount chosen.", "Parents apne bid par kitne yaqeen mein hain.", None),
        "wtWhy": ("Why not pay more?", "Main reason given by parents who bid below the maximum.", "Jinhon ne Rs 2,000 se kam bid kiya unhon ne ziyada kyun nahi diya.", None),
        "wtPractice": ("Practice round", "A Rs 0–20 practice before the real question.", "Asli sawal se pehle Rs 0 se 20 ki mashq (practice). Yeh sirf samajhne ke liye thi.", None),
        "wtSel": ("Understanding check 1", "If selected and the random price is above the bid, is the Rs 1,000 bonus still received?", "Samajh ki jaanch: kya parents ne bonus wala rule sahi samjha.", None),
        "wtNot": ("Understanding check 2", "If not selected, which bid is used?", "Samajh ki jaanch: agar parent muntakhib na ho to kaun sa bid use hota hai.", None),
        # ---- effects
        "teDemand": ("Did the videos change how many parents would pay?", "One line per group. A higher line means parents value the purifier more.",
                     "Teeno groups ki demand curve ek saath. Jo line upar hai us group ke parents ziyada dene ko tayyar hain. Lines qareeb hon to video ka koi wazeh asar nahi.", "demand"),
        "teArmCi": ("Average amount parents would pay, by group", "Bar = average bid. Thin line = the likely range.",
                    "Control, Video 1 aur Video 2 ka average bid. Patli line mumkin range (confidence interval) hai. Agar do groups ki lines ek doosre par chadhi hon to farq wazeh nahi.", "ci"),
        "teEff": ("Did the videos change what parents say?", "Each row compares two groups. Switch the comparison with the buttons.",
                  "Har row do groups ka muqabla hai. 'Difference' = video walay group minus control. 'Likely range' woh range hai jis mein asal farq hone ka imkaan hai. p-value batati hai ke yeh farq sirf ittefaq ho sakta hai ya nahi: 0.05 se chhoti ho to aam tor par asli farq maana jata hai. Aakhri column seedha lafzon mein nateeja likhta hai.", "effects"),
        "teBal": ("Are the three groups similar?", "A fairness check before comparing them.", "Muqabla tabhi fair hai jab teeno groups shuru mein ek jaise hon (umar, aamdani, area). ⚑ ka matlab hai ek characteristic mein farq nazar aaya, dobara dekhna chahiye.", "balance"),
        "teOrder": ("Does asking about other parents first change the bid?", "Average bid when parents were asked about other parents before or after their own bid.",
                    "Kuch parents se doosron ke bare mein pehle poocha gaya aur kuch se bid ke baad. Yeh dekhne ke liye ke sawal ka silsila jawab badalta hai ya nahi.", "ci"),
        # ---- fairness
        "faAgree": ("Who should pay for classroom air purifiers?", "How much parents agree with each statement.", "Teen bayanaat par parents kitne muttafiq hain. Gehra neela = zor se agree, laal = disagree.", "likert"),
        "faFree": ("Would parents let others pay?", "Share answering yes.", "Kya parents doosron ke pay karne par khud pay karna chhor denge (free-riding).", None),
        "faKnow": ("How many other parents do they know?", "Parents in the same classroom known personally.", "Class ke kitne doosre parents ko jaante hain.", None),
        "faPrio": ("Where does air quality rank among government priorities?", "Share who ranked each area first or second.", "Hukumat ki tarjeehat mein parents ne kis cheez ko pehle ya doosre number par rakha. Lambi bar = ziyada ahem.", None),
        "faNgo": ("Which charity would parents pick?", "If given Rs 3,000 to donate.", "Rs 3,000 dene hon to kis idare ko denge.", None),
        "faPol": ("Would parents back a tax-funded purifier plan?", "Purifiers in all primary classrooms with partial funding from taxes.", "Kya parents tax se purifier lagane ki policy ka saath denge.", None),
        "faSchool": ("What should schools do?", "Share answering yes.", "School ki fee, learning aur tarjeeh ke bare mein haan ka %.", None),
        # ---- time
        "tpDisc": ("How much extra do parents want to wait a year?", "The yearly return at which parents switch from money now to more money later.", "Parents Rs 2,000 abhi lena chahtay hain ya saal baad ziyada. Bar batati hai kitna ziyada mangte hain (per year). Bohat ziyada % = bohat jaldi paisa chahiye.", "discount"),
        "tpQual": ("How patient do parents say they are?", "0 = not willing to give up something today, 10 = very willing.", "Parents khud ko kitna sabr wala kehte hain (0 se 10).", None),
        "tpFirst": ("Rs 2,000 now or more in a year?", "The opening question of the money-now-or-later game.", "Pehla sawal: Rs 2,000 abhi ya taqreeban Rs 4,100 saal baad.", None),
        "tpCheck": ("Consistency check", "Rs 2,000 today or the same Rs 2,000 in one year?", "Jaanch: Rs 2,000 abhi ya wahi Rs 2,000 saal baad. Aksar 'abhi' chunna chahiye. Warna shak hai ke sawal samjha nahi.", None),
        "tpNet": ("Which mobile network do parents use?", "Network for the top-up thank-you.", "Top-up bhejne ke liye parents ka mobile network.", None),
        "tpTop": ("Top-up amounts", "Rs 2,000 less what the parent paid.", "Har parent ko Rs 2,000 mein se jo unhon ne contribute kiya woh ghata kar top-up mila.", None),
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
                nm = lambda r: f"{r['g']}: {r['l']}" if r["g"] not in ("Study arm",) else r["l"]
                return (f"Highest average: <strong>{nm(hi)}</strong> (Rs {hi['m']:,.0f}). Lowest: <strong>{nm(lo)}</strong> (Rs {lo['m']:,.0f}). "
                        "Where two thin lines overlap a lot, those groups are not clearly different.")
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
