

I wanted to share something I've recently put together that I believe will make a difference and save our team a meaningful amount of time and effort spent on manual data entry — every exam cycle.

**What I built**
A internal tool that does two things we've always done by hand:

1. **Loads every student's exam answers into our system automatically.** For the last paper alone, that's over 4,000 entries that used to be typed in one row at a time — now done in seconds, with no manual typing. Every paper going forward gets the same benefit.
2. **A visual analytics dashboard** — one click turns the same marks Excel into around 25 clear graphs across six tabs, all in plain language, each chart explaining itself.

**Why it matters — what we can now see, in plain words**

*Overview*
- **How the class spread overall** — a histogram showing how many students scored high, middling, or low. One glance tells you whether the class did well or struggled.
- **Which skills weigh most in the paper** — a donut showing how the 60 marks split across Remember / Understand / Apply / Analyze / Evaluate. Reveals which skills the exam actually weighs heaviest.

*Competency (skill-wise)*
- **Which skill the class finds easy or hard** — average % per skill. Tallest bar = easiest skill; shortest = hardest. Immediately shows where to focus teaching.
- **How many students totally missed a skill** — e.g. "8 students scored 0 in every Analyze question." Tall bar = a skill widely not understood.
- **How many students mastered a skill** — e.g. "15 students got full marks in every Remember question." Tall bar = broadly mastered.
- **Exactly which students** scored 0 or maxed each skill — a table that names the students behind each count, so support can be targeted by name, not just numbers.

*Questions*
- **Which questions are hard** — average % per question. Low bars flag the questions to review in class.
- **Which questions students skipped** — attempt rate per question. A dip = a question many left blank or got fully wrong.
- **What to re-teach and what's working** — Hardest 5 and Easiest 5 tables, ranked, so you can act on the few questions that matter most.

*Students*
- **Every student ranked best-to-worst** — a clear ranking with hover showing each student's name, section, and total %. No more guessing who's where.
- **Top 10 toppers and Bottom 10 strugglers** — tables broken down by skill, so you can see how each topper got there, and which skill each struggling student most needs help with.
- **Exactly which students are in each band** — a table naming the students in Need Improvement / Adequate / Sufficient / Excellent, with their total %, ready to pull into a support list.

*Sections*
- **Which section leads overall** — average total % by section. Tallest bar = the best-performing section.
- **Which skill each section is strong or weak in** — grouped bars: within each skill, compare every section. Shows exactly where each section's strengths and gaps are.
- **Which section has the most students below pass** — pass rate by section. A short bar = a section needing attention.

*Performance bands*
- **How the class splits overall** — a donut showing the share of students in Need Improvement / Adequate / Sufficient / Excellent. Big red slice = many need support.
- **How each section splits** — stacked bars per section. A section heavy on red = needs the most attention.
- **A check that the dashboard matches our Excel** — a table showing the dashboard's band counts next to the Excel's own counts, so we can confirm they always agree.

*Bonus (optional): generate the SQL in the browser*
- For anyone who needs to load the answers but doesn't have direct database access, the dashboard can produce the INSERT statements as text to copy — gated by a small file that only a DB-access user can produce, so graph viewers can't, only our data operator can.

**The bigger picture**
- **Time saved:** thousands of manual entries per exam, removed — time back for the team every paper cycle.
- **Sharper decisions:** leadership gets a clean, presentable view instead of spreadsheets — which skill to focus on, which questions to re-teach, which students to support, and how each section is doing.
- **Consistent and accurate:** the data is validated against the system, so what's shown is exactly what's stored — no manual slips.
- **Ready for any exam:** it adapts to whatever paper structure we use next — nothing needs rebuilding when the format changes.
- **Private and secure:** the graphs never touch our system's data; the loading runs only where it should. Nothing sensitive leaves our machines.

What used to be a slow, manual, repeat-every-cycle task is now a one-step process, and the same marks give us a far richer picture to act on — for teachers, for section heads, and for leadership.

I'd be glad to walk you through it in 10 minutes and show the difference on a real paper.

Best regards,
Rohit Kumar