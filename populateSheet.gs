// Paste this into Google Apps Script — replace ALL code in Code.gs, delete New.gs entirely.
// Two functions available:
//   populateIssues()   → run this first to create/update the Issues tab
//   populateAllRaces() → run this to create/update the Statewide candidates tab

function populateIssues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Issues") || ss.insertSheet("Issues");
  sheet.clearContents();

  var headers = [
    "id","topic","question",
    "optionA","shortA",
    "optionB","shortB",
    "optionC","shortC",
    "optionD","fact"
  ];

  var NEUTRAL = "This issue is important, but it won't decide my vote.";

  var issues = [
    ["reproductiveRights",
     "Reproductive Rights",
     "When it comes to reproductive healthcare decisions, I believe:",
     "Pregnancy and abortion decisions belong between a person and their doctor — not politicians. Texas's current total ban goes too far and puts lives at risk.",
     "Abortion decisions belong to patients and their doctors — not politicians.",
     "I believe abortion is a deeply personal decision, and I understand the moral weight on both sides. I support access to abortion in early pregnancy with reasonable limits, and I want clear exceptions that protect women's health and lives in all circumstances.",
     "Abortion should be legal in early pregnancy with clear exceptions for health and safety.",
     "Life begins at conception, and government has a responsibility to protect the unborn. Texas's abortion laws reflect our communities' deeply held values.",
     "Life begins at conception; government must protect the unborn.",
     NEUTRAL,
     "Texas enacted a near-total abortion ban in 2022 after the Supreme Court overturned Roe v. Wade. Exceptions exist only for life-threatening medical emergencies."],

    ["immigration",
     "Immigration & Border",
     "On immigration and the Texas border, I believe:",
     "Immigrants strengthen Texas economically and culturally. We need clear, humane legal pathways to citizenship — not mass deportation or family separation.",
     "Immigrants strengthen Texas; we need humane, legal pathways to citizenship.",
     "The border needs to be secure — but mass deportation isn't the answer. I want enforcement that's firm but humane, and a realistic legal pathway for the millions of people who've built their lives and families here.",
     "We need secure borders AND a humane, realistic path for people already here.",
     "Securing the border is the top priority. Illegal immigration strains public resources and must be stopped through strong enforcement.",
     "Securing the border is the top priority — stop illegal crossings.",
     NEUTRAL,
     "Texas shares a 1,254-mile border with Mexico. Illegal border crossings hit a record high in 2024 before declining sharply in early 2025 after federal policy changes."],

    ["gunPolicy",
     "Gun Policy",
     "On gun laws, I believe:",
     "Common-sense measures — universal background checks, red flag laws, and assault weapon restrictions — save lives without infringing on law-abiding citizens' rights.",
     "Background checks and red flag laws save lives without banning guns.",
     "I believe in the Second Amendment and the right to own firearms. I also think background checks and keeping guns out of dangerous hands are reasonable and don't threaten law-abiding gun owners. Rights come with responsibilities.",
     "I support gun rights and think basic safety measures like background checks are reasonable.",
     "The Second Amendment is a fundamental right that must be fully protected. Government gun restrictions do more harm than good and leave people less safe.",
     "The Second Amendment must be fully protected from any restriction.",
     NEUTRAL,
     "Texas passed permitless carry in 2021. Firearms were involved in over 4,100 Texas deaths in 2022. Texas has more gun dealers than any other state."],

    ["education",
     "Public Education",
     "On public schools and education funding, I believe:",
     "Tax dollars belong in public schools. Voucher programs that send public money to private or religious schools undermine the communities that depend on them.",
     "Tax dollars belong in public schools — not private school vouchers.",
     "Strong public schools are the backbone of our communities and deserve full funding. I also think parents should have some meaningful say in their child's education — as long as accountability follows the money, wherever it goes.",
     "Fund public schools fully, but give parents some meaningful choice with accountability.",
     "Parents — not government — should decide where their child is educated. Funding should follow the student to whatever school best meets their needs.",
     "Parents, not government, should choose their child's school.",
     NEUTRAL,
     "Texas funds 1,200+ school districts. In 2023 the legislature passed a $1B school voucher program; opponents say it diverts funds from underfunded public schools."],

    ["healthcare",
     "Healthcare",
     "On healthcare access for Texans, I believe:",
     "Every Texan deserves access to affordable healthcare. Texas should expand Medicaid, protect the ACA, and ensure no one goes without coverage due to cost.",
     "Every Texan deserves affordable coverage — expand Medicaid now.",
     "Healthcare costs are crushing Texas families and something has to change. I'm open to both market-based reforms and expanded public programs — whatever actually lowers costs and gets people covered without breaking the budget.",
     "Healthcare needs reform — I'm open to market and public solutions that actually lower costs.",
     "Healthcare works best when driven by free-market competition, not government programs. Expanding Medicaid grows dependency and increases costs for taxpayers.",
     "Free-market competition, not government programs, lowers healthcare costs.",
     NEUTRAL,
     "Texas has the highest uninsured rate in the nation at 18.3%. It's one of only 10 states that has not expanded Medicaid under the Affordable Care Act."],

    ["climate",
     "Climate & Energy",
     "On energy and the environment, I believe:",
     "Texas must invest in clean, renewable energy and hold polluters accountable. The ERCOT crisis showed what happens when we prioritize profit over reliability and sustainability.",
     "Invest in clean energy and hold polluters accountable.",
     "Texas's energy industry matters and so does our environment. We can invest in cleaner, more reliable energy while protecting jobs — the 2021 grid failure showed that the status quo isn't working for anyone.",
     "We need reliable energy AND a cleaner future — balance both without sacrificing either.",
     "Texas's oil and gas industry drives our economy and national energy security. Over-regulation threatens jobs and energy independence without meaningful climate benefit.",
     "Oil and gas drive Texas prosperity — protect them from over-regulation.",
     NEUTRAL,
     "Texas is the nation's top producer of both oil and wind energy. The 2021 winter storm knocked out power for 4.5 million homes and caused $130B+ in damages."],

    ["democracy",
     "Money in Politics",
     "On money, corruption, and political influence, I believe:",
     "Dark money and corporate PACs have corrupted our politics. We need strict campaign finance limits, full donor transparency, and an end to politicians being bought by special interests.",
     "Dark money and corporate PACs have corrupted politics — we need strict limits.",
     "There's too much money influencing our politics — that's not a partisan opinion, it's what most Texans believe. I'd support full donor transparency and disclosure requirements as a starting point, even if I'm cautious about limiting political speech outright.",
     "Support full donor transparency; cautious about limiting political speech outright.",
     "Political donations are free speech. The problem isn't money in politics — it's government that's too big and too powerful to resist buying. Limit government, not speech.",
     "Political donations are free speech. Limit government power, not political spending.",
     NEUTRAL,
     "74% of Americans across both parties say money in politics is a major problem. Texas lawmakers received over $200M in PAC donations in the 2024 election cycle."],

    ["socialSecurity",
     "Social Security & Medicare",
     "On Social Security and Medicare, I believe:",
     "These programs must be fully protected and strengthened. Seniors earned these benefits — any cuts or privatization are unacceptable.",
     "Protect and strengthen Social Security and Medicare — no cuts ever.",
     "These programs earned the trust of millions of Texans and shouldn't be gutted or privatized. If honest, gradual adjustments are needed to keep them strong for future generations, I'm open to that conversation — but only if benefits are protected.",
     "Protect these programs — open to gradual adjustments only if benefits stay intact.",
     "These programs face long-term insolvency and need structural reform. Responsible changes now will ensure future generations can still benefit.",
     "Long-term solvency requires honest structural reform starting now.",
     NEUTRAL,
     "4.7 million Texans receive Social Security or Medicare benefits. Without legislative changes, the Social Security trust fund is projected to be depleted by 2033."],

    ["affordability",
     "Housing & Affordability",
     "On the cost of housing and everyday life in Texas, I believe:",
     "Texas must invest in affordable housing, protect renters, and reform zoning laws. Rising costs are pushing working families out of their communities.",
     "Protect renters and invest in affordable housing — the market isn't working.",
     "Rising costs are pushing working families out of the communities they built. We need more housing supply AND some basic protections for renters — the free market alone hasn't solved it, but heavy-handed government intervention has its own problems.",
     "We need more housing supply AND basic renter protections — both matter.",
     "Government intervention inflates housing costs. Deregulation, reduced permitting barriers, and free-market competition are the fastest path to affordable homes.",
     "Deregulation and free-market competition are the real path to affordability.",
     NEUTRAL,
     "Texas median home prices rose over 50% from 2020–2023. Austin and San Marcos now rank among the least affordable markets in the South for median-income households."]
  ];

  var allRows = [headers].concat(issues);
  sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1e293b");
  headerRange.setFontColor("#ffffff");

  // Color-code columns by position type
  var colors = {
    3: "#dbeafe", 4: "#dbeafe",   // optionA, shortA — blue
    5: "#ede9fe", 6: "#ede9fe",   // optionB, shortB — purple
    7: "#fee2e2", 8: "#fee2e2",   // optionC, shortC — red
    9: "#f3f4f6",                 // optionD — gray
    10: "#f0fdf4"                 // fact — green-tint
  };
  for (var col in colors) {
    sheet.getRange(2, parseInt(col), issues.length, 1).setBackground(colors[col]);
  }

  // Column widths
  var widths = [20,18,35,45,30,45,30,45,30,35,45];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i+1, w*6); });

  // Wrap text
  sheet.getRange(1, 1, allRows.length, headers.length).setWrap(true);

  Logger.log("SUCCESS: Issues tab populated with " + issues.length + " issues (4 positions each).");
}

// ─────────────────────────────────────────────────────────────────────────────

function populateAllRaces() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var stSheet = ss.getSheetByName("Statewide") || ss.getSheetByName("statewide");
  if (!stSheet) { stSheet = ss.insertSheet("Statewide"); }
  stSheet.clearContents();

  var A = "A";
  var B = "B";

  var headers = ["office","category","districtType","districtNumber","name","party","confirmed","description","summary","website","photo","reproductiveRights","immigration","gunPolicy","education","healthcare","religionAndLiberty","climate","democracy","socialSecurity","water"];

  // Drive photo base URL
  var D = "https://lh3.googleusercontent.com/d/";

  // Each entry: [office, category, districtType, districtNumber, description, name, party, confirmed, summary, photo]
  var candidates = [

    // ── FEDERAL ──────────────────────────────────────────────────────────────
    ["U.S. Senate","federal","","",
     "The U.S. Senator represents all Texans in Washington for a 6-year term, voting on federal legislation, Supreme Court nominees, and treaties.",
     "James Talarico","D","true",
     "Former state rep turned Senate candidate. Progressive Democrat supporting abortion rights, Medicaid expansion, public education funding, and gun safety reform.",
     D+"1xT_tzvyzpedvvRDZJ3V09TqjcoTiPiDO"],
    ["U.S. Senate","federal","","",
     "The U.S. Senator represents all Texans in Washington for a 6-year term, voting on federal legislation, Supreme Court nominees, and treaties.",
     "Ken Paxton","R","true",
     "Former Texas AG acquitted after 2023 impeachment. Supports border enforcement, opposes abortion rights and gun restrictions. Endorsed by President Trump.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/2/20/K_Paxton.jpg/250px-K_Paxton.jpg"],

    ["U.S. House — District 21","federal","congressional","21",
     "Covers western Hays County including Wimberley and parts of San Marcos.",
     "Kristin Hook","D","true",
     "Democrat advancing from the March 2026 primary for Texas' 21st Congressional District. Supports reproductive rights, climate action, and working families.",
     D+"1VSck3tC9d9U4xX3sHPxUjD7QVvdl63cf"],
    ["U.S. House — District 21","federal","congressional","21",
     "Covers western Hays County including Wimberley and parts of San Marcos.",
     "Mark Teixeira","R","true",
     "Republican nominee for CD-21. Won the March 2026 primary with 62% of the vote. Conservative on border security, taxes, and government spending.",
     ""],

    ["U.S. House — District 27","federal","congressional","27",
     "Covers eastern Hays County including Kyle and Buda. Incumbent Michael Cloud (R) seeks a fourth term.",
     "Tanya Lloyd","D","true",
     "Democrat who defeated two opponents in the March 2026 primary for Texas' 27th Congressional District. Challenger to incumbent Michael Cloud.",
     D+"1d7uIYEn5PTzGI-IhOcBF2NNwCvZBbqXx"],
    ["U.S. House — District 27","federal","congressional","27",
     "Covers eastern Hays County including Kyle and Buda.",
     "Michael Cloud","R","true",
     "Three-term incumbent Congressman representing CD-27. Strong conservative on border security, Second Amendment, and limited government.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/2/2d/Michael_Cloud%2C_Official_Portrait%2C_115th_Congress.jpg/250px-Michael_Cloud%2C_Official_Portrait%2C_115th_Congress.jpg"],

    // ── STATE ─────────────────────────────────────────────────────────────────
    ["Governor of Texas","state","","",
     "The Governor leads state government, signs or vetoes legislation, and sets the policy agenda for Texas.",
     "Gina Hinojosa","D","true",
     "Five-term Austin state rep and civil rights attorney who won the Democratic primary with 59%. Running against Abbott's school voucher program and corporate influence in state government.",
     D+"1Vgf4PVr3BTXCLJk9xcKh5jzf0HDUWEVv"],
    ["Governor of Texas","state","","",
     "The Governor leads state government, signs or vetoes legislation, and sets the policy agenda for Texas.",
     "Greg Abbott","R","true",
     "Three-term incumbent Governor. Signed Texas's abortion ban, constitutional carry, and school voucher legislation. Launched Operation Lone Star border initiative.",
     "https://gov.texas.gov/uploads/images/general/2024-GovernorAbbott-Portrait.jpg"],

    ["Lieutenant Governor","state","","",
     "The Lt. Governor runs the Texas Senate and shapes the state budget. Arguably the most powerful office in Texas government.",
     "Vikki Goodwin","D","true",
     "Austin state representative who won the Democratic nomination for Lt. Governor. Supports public education funding, reproductive rights, and healthcare access.",
     D+"1bR4y9O7X7ngfL1CWfKZOap5OZuVpSCSe"],
    ["Lieutenant Governor","state","","",
     "The Lt. Governor runs the Texas Senate and shapes the state budget. Arguably the most powerful office in Texas government.",
     "Dan Patrick","R","true",
     "Three-term incumbent Lt. Governor. Architect of Texas's abortion ban, school voucher push, anti-trans legislation, and SB 1 voting restrictions.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/4/41/Dan_Patrick_Texas_%28cropped%29.jpg/250px-Dan_Patrick_Texas_%28cropped%29.jpg"],

    ["Attorney General","state","","",
     "Texas's chief legal officer, enforcing state laws on consumer protection, environmental rules, voting rights, and more.",
     "Nathan Johnson","D","true",
     "Dallas State Senator who defeated Joe Jaworski in the Democratic runoff. Former judge focused on environmental enforcement, voting rights, and reproductive freedom.",
     D+"1RsDuQh5hf6K5At8DwF1NzmAhIEob7iG_"],
    ["Attorney General","state","","",
     "Texas's chief legal officer, enforcing state laws on consumer protection, environmental rules, voting rights, and more.",
     "Mayes Middleton","R","true",
     "State Senator and oil & gas executive who defeated Congressman Chip Roy in the May 26 runoff. One of the most conservative members of the Texas Senate. Self-funded nearly $17M.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/d/de/Mayes_Middleton_by_Gage_Skidmore.jpg/250px-Mayes_Middleton_by_Gage_Skidmore.jpg"],

    ["Comptroller of Public Accounts","state","","",
     "Texas's chief financial officer — collects state taxes, estimates revenue for the state budget, and manages state funds.",
     "Sarah Eckhardt","D","true",
     "Former Travis County Judge and state senator. Democrat challenging for Comptroller with a focus on fiscal transparency, clean energy investment, and sustainable growth.",
     D+"1IZ44ntovwR7zr1Xi87H3V2ex37aoW5w_"],
    ["Comptroller of Public Accounts","state","","",
     "Texas's chief financial officer — collects state taxes, estimates revenue for the state budget, and manages state funds.",
     "Don Huffines","R","true",
     "Republican nominee for Comptroller. Former state senator who ran for governor in 2022. Staunch fiscal conservative focused on cutting taxes and government spending.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/Don_Huffines_%2851370822589%29_%28cropped%29.jpg/250px-Don_Huffines_%2851370822589%29_%28cropped%29.jpg"],

    ["Commissioner of General Land Office","state","","",
     "Manages 13 million acres of state land, the Alamo, and disaster recovery funds. Oversees the Permanent School Fund.",
     "Benjamin Flores","D","true",
     "Democratic nominee for Texas Land Commissioner. Running against incumbent Dawn Buckingham on a platform of transparent management of state lands and disaster recovery funds.",
     D+"1tFOE9x4zS9ZMS9KXIpvcRQLB3IHuVzX7"],
    ["Commissioner of General Land Office","state","","",
     "Manages 13 million acres of state land, the Alamo, and disaster recovery funds. Oversees the Permanent School Fund.",
     "Dawn Buckingham","R","true",
     "Incumbent Land Commissioner elected in 2022. Former State Senator focused on school funding, veterans programs, and management of state natural resources.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d6/Sen._Dawn_Buckingham%2C_M.D_%28cropped%29.jpg/250px-Sen._Dawn_Buckingham%2C_M.D_%28cropped%29.jpg"],

    ["Commissioner of Agriculture","state","","",
     "Regulates food safety, agriculture, and rural development across Texas — one of the top agricultural states in the country.",
     "Clayton Tucker","D","true",
     "Democratic nominee for Texas Agriculture Commissioner. Focused on supporting family farmers, food safety standards, and rural community investment.",
     D+"1TPfhGdxTp1tuOQ_FlzzIne09bgs6H7Yh"],
    ["Commissioner of Agriculture","state","","",
     "Regulates food safety, agriculture, and rural development across Texas — one of the top agricultural states in the country.",
     "Nate Sheets","R","true",
     "Republican nominee who ousted three-term incumbent Sid Miller in the March 2026 primary. Focused on deregulation, water rights, and supporting Texas agriculture.",
     ""],

    ["Railroad Commissioner","state","","",
     "Despite the name, regulates the oil, gas, and pipeline industries — one of the most powerful economic offices in Texas.",
     "Jon Rosenthal","D","true",
     "Democratic nominee for Railroad Commission. Former state rep focused on pipeline safety, grid reliability after Winter Storm Uri, and energy transition.",
     D+"1P6JDF1g-eNMTWBfLJsySpjGg8hBITIwq"],
    ["Railroad Commissioner","state","","",
     "Despite the name, regulates the oil, gas, and pipeline industries — one of the most powerful economic offices in Texas.",
     "Bo French","R","true",
     "Republican nominee who defeated incumbent Jim Wright in the May 26 runoff. Oil and gas industry background; supports expanded fossil fuel production.",
     "https://bofrench.com/wp-content/uploads/2025/11/BoFrench05-updated.jpg"],

    ["State Board of Education, District 5","state","","",
     "Sets curriculum standards, approves textbooks, and oversees funding for Texas public schools.",
     "Allison Bush","D","true",
     "Democrat who won the May 26 runoff with 61% of the vote. Opposes vouchers and book bans, supports inclusive curriculum and public school funding.",
     D+"1xM-bnvTwQOhsYuZ5ignKNr6nAEzxb2fc"],
    ["State Board of Education, District 5","state","","",
     "Sets curriculum standards, approves textbooks, and oversees funding for Texas public schools.",
     "Mica Arellano","R","true",
     "Republican nominee for SBOE District 5. Aligned with GOP education priorities including school choice, parental rights in curriculum, and Texas history standards.",
     "https://assets.cdn.filesafe.space/ggxeWPIAhnIsXRqvLmwz/media/69f54c71bbacf0a05623b135.jpg"],

    ["TX State Senate — District 21","state","stateSenate","21",
     "The State Senator for District 21 serves Hays County in the Texas Senate, voting on state budget, legislation, and judicial confirmations.",
     "Amy Martinez-Salas","D","true",
     "Democratic nominee for Texas State Senate District 21. Running on public education, healthcare access, and responsible water management in the Hill Country.",
     ""],
    ["TX State Senate — District 21","state","stateSenate","21",
     "The State Senator for District 21 serves Hays County in the Texas Senate, voting on state budget, legislation, and judicial confirmations.",
     "David Cook","R","true",
     "Republican nominee for State Senate District 21. Won the March 2026 primary. Conservative focused on property tax relief, border security, and Second Amendment rights.",
     "https://www.house.texas.gov/images/members/3960.jpg"],

    ["TX State House — District 45","state","stateHouse","45",
     "Covers parts of Hays County including Kyle and portions of San Marcos. The State Representative votes on Texas law and the state budget.",
     "Erin Zwiener","D","true",
     "Incumbent State Representative for HD-45. Environmental scientist and progressive Democrat who has championed water policy, reproductive rights, and public school funding since 2019.",
     D+"1u_3_4AExJvyhg_bl5uTPDqYvhRCa8r1X"],
    ["TX State House — District 45","state","stateHouse","45",
     "Covers parts of Hays County including Kyle and portions of San Marcos. The State Representative votes on Texas law and the state budget.",
     "Tennyson Moreno","R","true",
     "Republican nominee for HD-45. Challenging incumbent Erin Zwiener in this competitive Hays County district. Focuses on property taxes, parental rights, and border enforcement.",
     ""],

    ["TX State House — District 73","state","stateHouse","73",
     "Covers parts of Hays County including Wimberley and Dripping Springs. The State Representative votes on Texas law and the state budget.",
     "Merrie Fox","D","true",
     "Democratic nominee for HD-73. Challenger to Republican incumbent Carrie Isaac in this Hill Country district. Focused on water rights, public education, and reproductive freedom.",
     D+"1Z3AqBu6kSeJoXY0E-M0QZAfYan1jisF7"],
    ["TX State House — District 73","state","stateHouse","73",
     "Covers parts of Hays County including Wimberley and Dripping Springs. The State Representative votes on Texas law and the state budget.",
     "Carrie Isaac","R","true",
     "Republican incumbent State Representative for HD-73. Aligned with state Republican priorities on school choice, property taxes, and parental rights in education.",
     "https://www.house.texas.gov/images/members/4265.jpg?v=3"],

    // ── COUNTY ───────────────────────────────────────────────────────────────
    ["Hays County Judge","county","","",
     "The County Judge leads the Commissioners Court, overseeing the county budget, emergency management, and local policy.",
     "Michelle Gutierrez Cohen","D","true",
     "Former Hays County Commissioner Pct. 2 who defeated incumbent Ruben Becerra in the May 26 runoff with 58% of the vote. Focused on infrastructure, growth management, and county services.",
     D+"1FWXoq0slIB7RjQmKcbrERzq2JSfLKlJS"],
    ["Hays County Judge","county","","",
     "The County Judge leads the Commissioners Court, overseeing the county budget, emergency management, and local policy.",
     "Geoffrey Tahuahua","R","true",
     "Republican nominee for Hays County Judge. Running on fiscal conservatism, public safety, and managing the county's rapid growth responsibly.",
     "https://images.squarespace-cdn.com/content/v1/67e5a3c13449f136b69128f2/3a6941ac-1390-4d40-8e31-5bf709dbcb03/CE6A1ECA-3AFC-421A-9C4A-FE26FC053D6C.JPG"],

    ["Criminal District Attorney","county","","",
     "Prosecutes criminal cases in Hays County, including misdemeanors and felonies. Sets prosecution priorities and manages the DA's office.",
     "Alfonso Salazar","D","true",
     "Won the Democratic primary with 54% — no Republican filed, making this race effectively decided. Military veteran who served in the Air National Guard and on active duty. Challenger to the DA's office status quo.",
     D+"1WEmZ-faHrCZ5DX4mfS4rgnUVFWv7NfME"],

    ["District Clerk","county","","",
     "Maintains court records, processes legal filings, and manages jury pools for the district courts of Hays County.",
     "Amanda Calvert","D","true",
     "Democratic nominee for Hays County District Clerk. Focused on accessible court records and efficient administration of the district court system.",
     D+"15dgiRVnmu4Dm6YBjuFC3-HFRKHUeJowx"],
    ["District Clerk","county","","",
     "Maintains court records, processes legal filings, and manages jury pools for the district courts of Hays County.",
     "TBD — Republican Nominee","R","false",
     "Republican nominee to be confirmed.",
     ""],

    ["County Clerk","county","","",
     "Manages elections, vital records, real estate documents, and county court records for Hays County.",
     "Elaine H. Cárdenas","D","true",
     "Democratic nominee for Hays County Clerk. Focused on election administration, accessible public records, and efficient county services.",
     ""],
    ["County Clerk","county","","",
     "Manages elections, vital records, real estate documents, and county court records for Hays County.",
     "TBD — Republican Nominee","R","false",
     "Republican nominee to be confirmed.",
     ""],

    ["County Treasurer","county","","",
     "Receives, deposits, and disburses county funds. Manages Hays County's financial accounts and investment portfolio.",
     "Daphne Tenorio","D","true",
     "Democratic nominee for Hays County Treasurer. Focused on responsible management of county funds and transparent financial reporting.",
     D+"1WBbkM_X7xCKkj7WIGaPi9nQqbILnP8xl"],
    ["County Treasurer","county","","",
     "Receives, deposits, and disburses county funds. Manages Hays County's financial accounts and investment portfolio.",
     "TBD — Republican Nominee","R","false",
     "Republican nominee to be confirmed.",
     ""],

    ["Commissioner, Pct 2","county","commissioner","2",
     "Hays County Commissioner for Precinct 2. Sits on the Commissioners Court, voting on the county budget, roads, and local services.",
     "Johnny Flores","D","true",
     "Hays CISD trustee and Democratic nominee for Commissioner Pct. 2. Focused on education, infrastructure, and responsible development in fast-growing central Hays County.",
     D+"1zCPZ03n9VzfJb9kqf5lTMsBdADoLv1TD"],
    ["Commissioner, Pct 2","county","commissioner","2",
     "Hays County Commissioner for Precinct 2. Sits on the Commissioners Court, voting on the county budget, roads, and local services.",
     "Abby Gibson","R","true",
     "Republican nominee for Commissioner Precinct 2. Focused on fiscal responsibility, public safety, and managing growth in the Pct. 2 area.",
     ""],

    ["Commissioner, Pct 4","county","commissioner","4",
     "Hays County Commissioner for Precinct 4. Sits on the Commissioners Court, voting on the county budget, roads, and local services.",
     "Angie Unger","D","true",
     "Democratic nominee for Commissioner Precinct 4. Focused on infrastructure, environmental protection, and community services in the Dripping Springs area.",
     D+"16MDs2_MH6WBy1n9puhTtyLSEzPlCPawG"],
    ["Commissioner, Pct 4","county","commissioner","4",
     "Hays County Commissioner for Precinct 4. Sits on the Commissioners Court, voting on the county budget, roads, and local services.",
     "Rob McClelland","R","true",
     "Republican nominee for Commissioner Pct. 4. Won the March 2026 Republican primary. Focused on fiscal conservatism, road infrastructure, and controlled growth.",
     "https://lirp.cdn-website.com/3ea4577a/dms3rep/multi/opt/McClelland-Pic-1920w.png"],

    ["Justice of Peace, Pct 1, Pl 2","county","jpPrecinct","1",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 1.",
     "TBD — Democratic Nominee","D","false","Democratic nominee to be confirmed.",""],
    ["Justice of Peace, Pct 1, Pl 2","county","jpPrecinct","1",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 1.",
     "TBD — Republican Nominee","R","false","Republican nominee to be confirmed.",""],

    ["Justice of Peace, Pct 2, Pl 1","county","jpPrecinct","2",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 2.",
     "TBD — Democratic Nominee","D","false","Democratic nominee to be confirmed.",""],
    ["Justice of Peace, Pct 2, Pl 1","county","jpPrecinct","2",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 2.",
     "TBD — Republican Nominee","R","false","Republican nominee to be confirmed.",""],

    ["Justice of Peace, Pct 5","county","jpPrecinct","5",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 5.",
     "TBD — Democratic Nominee","D","false","Democratic nominee to be confirmed.",""],
    ["Justice of Peace, Pct 5","county","jpPrecinct","5",
     "Handles small claims, Class C misdemeanors, and civil cases under $20,000 for JP Precinct 5.",
     "TBD — Republican Nominee","R","false","Republican nominee to be confirmed.",""],

    // ── JUDICIAL ─────────────────────────────────────────────────────────────
    ["TX Supreme Court — Chief Justice","judicial","","",
     "Leads the Texas Supreme Court — the highest civil court in Texas, deciding appeals on civil and family law matters statewide.",
     "Maggie Ellis","D","true",
     "Democratic nominee for Texas Supreme Court Chief Justice. Challenging incumbent Jimmy Blacklock in this statewide judicial race.",
     D+"1v5I1GRK4wPi4MKLwDDA1Th07tAHas-zK"],
    ["TX Supreme Court — Chief Justice","judicial","","",
     "Leads the Texas Supreme Court — the highest civil court in Texas, deciding appeals on civil and family law matters statewide.",
     "Jimmy Blacklock","R","true",
     "Incumbent Texas Supreme Court Chief Justice. Conservative jurist appointed by Governor Abbott in 2018. Supported overturning Roe v. Wade's application in Texas.",
     "https://www.txcourts.gov/media/1460030/chief-justice-jimmy-blacklock-web.jpg"],

    ["TX Supreme Court — Place 2","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "Chari Kelly","D","true",
     "Democratic nominee for Texas Supreme Court Place 2. Challenging incumbent James Sullivan in this statewide civil court seat.",
     D+"1mlfpbv16J2EVRaYMConG3igJQRa1f_RF"],
    ["TX Supreme Court — Place 2","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "James Sullivan","R","true",
     "Incumbent Texas Supreme Court Justice, Place 2. Conservative jurist appointed to the state's highest civil court.",
     "https://www.txcourts.gov/media/1460010/justice-james-p-sullivan-2025-web.jpg"],

    ["TX Supreme Court — Place 7","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "Kristen Hawkins","D","true",
     "Democrat who presides over the 11th District Court in Harris County. Defeated Gordon Goodman in the Democratic primary for this Texas Supreme Court seat.",
     D+"15526sVwf6jQ38Z3VYKx3Vdg9H8IB_NKf"],
    ["TX Supreme Court — Place 7","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "Kyle Hawkins","R","true",
     "Republican nominee for Texas Supreme Court Place 7. Former Texas Solicitor General appointed to the court.",
     "https://www.txcourts.gov/media/1461570/justice-kyle-hawkins.jpg"],

    ["TX Supreme Court — Place 8","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "Gisela Triana","D","true",
     "Democratic nominee for Texas Supreme Court Place 8. Challenging incumbent Brett Busby in this competitive statewide judicial race.",
     D+"1HT4cmpIVS5NvsDz7O9JQ-w1bLPrn6vY1"],
    ["TX Supreme Court — Place 8","judicial","","",
     "Texas Supreme Court Justice — decides civil and family law appeals at the state's highest court.",
     "Brett Busby","R","true",
     "Incumbent Texas Supreme Court Justice, Place 8. Conservative jurist on the state's highest civil court.",
     "https://www.txcourts.gov/media/1461169/brett-busby.png"],

    ["Court of Criminal Appeals — Place 3","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "Okey Anyiam","D","true",
     "Democratic nominee for Court of Criminal Appeals Place 3. Challenging Paxton-backed Thomas Smith in this key criminal appeals seat.",
     D+"12NVyJQtoVYH2UkhYA7n3llPi9yL0nK5-"],
    ["Court of Criminal Appeals — Place 3","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "Thomas Smith","R","true",
     "Republican nominee who defeated Alison Fox in the May 26 GOP runoff. Assistant AG in Ken Paxton's office, endorsed by Paxton for this criminal appeals seat.",
     "https://texasjudges.org/wp-content/uploads/2025/10/Thomas_Smith_400x400.jpg"],

    ["Court of Criminal Appeals — Place 4","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "Audra Riley","D","true",
     "Democratic nominee for Court of Criminal Appeals Place 4. Criminal district court judge in Dallas County challenging Republican incumbent Kevin Yeary.",
     D+"1UhakKiOPblZvFq93YhjT8LZBrd4Pz8wy"],
    ["Court of Criminal Appeals — Place 4","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "Kevin Patrick Yeary","R","true",
     "Incumbent Court of Criminal Appeals Judge, Place 4. Experienced criminal jurist running for another term on Texas's highest criminal court.",
     "https://www.txcourts.gov/media/1462538/yeary.jpg"],

    ["Court of Criminal Appeals — Place 9","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "Holly Taylor","D","true",
     "Democratic nominee for Court of Criminal Appeals Place 9. Challenging Republican John Messinger in this statewide criminal appeals seat.",
     D+"1_9309T_hCY0a9jDadCqMiDHfC3JvlZpr"],
    ["Court of Criminal Appeals — Place 9","judicial","","",
     "Texas's highest criminal court — reviews criminal appeals and death penalty cases from across the state.",
     "John Messinger","R","true",
     "Republican nominee for CCA Place 9. Assistant state prosecuting attorney who defeated Jennifer Balido in the March 2026 primary.",
     "https://texasjudges.org/wp-content/uploads/2024/02/john-messinger.png"],

    ["15th Court of Appeals — Chief Justice","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "Jerry Zimmerer","D","true",
     "Democratic nominee for 15th Court Chief Justice. Former justice on the Houston-based 14th Court of Appeals, running against appointed Republican Scott Brister.",
     D+"1JLpsdgnKp0y6mzPpjEN3jmyM_xf_w23-"],
    ["15th Court of Appeals — Chief Justice","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "Scott Brister","R","true",
     "Republican Chief Justice of the 15th Court of Appeals, appointed by Governor Abbott when the court was created in 2024. Former Texas Supreme Court Justice.",
     "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e0/Brister_Scott_hiresRGB_12_2012.jpg/250px-Brister_Scott_hiresRGB_12_2012.jpg"],

    ["15th Court of Appeals — Place 2","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "Tom Baker","D","true",
     "Democratic nominee for 15th Court Place 2. Former justice on the 3rd Court of Appeals in Austin, challenging appointed Republican Scott Field.",
     D+"1KJQsFQyoqny_UEZFZIA7cyLdXkm4RyGE"],
    ["15th Court of Appeals — Place 2","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "Scott Field","R","true",
     "Republican Justice of the 15th Court of Appeals, Place 2, appointed by Governor Abbott when the court was created in 2024.",
     "https://www.txcourts.gov/media/1459003/justice-scott-field-bio-photo.jpg"],

    ["15th Court of Appeals — Place 3","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "Marc Meyer","D","true",
     "Democratic nominee for 15th Court Place 3. Private law firm attorney challenging appointed Republican April Farris.",
     D+"1G2gQVBFd_kQAdAwP-J-hV2UZZy96i4YJ"],
    ["15th Court of Appeals — Place 3","judicial","","",
     "A new statewide appellate court created by the Texas Legislature. Hears civil cases involving the state government.",
     "April Farris","R","true",
     "Republican Justice of the 15th Court of Appeals, Place 3, appointed by Governor Abbott when the court was created in 2024.",
     "https://www.txcourts.gov/media/1461250/af-web-page-091025.jpg"],

    ["3rd Court of Appeals — Chief Justice","judicial","","",
     "The 3rd Court of Appeals serves Central Texas including Hays County — hears civil and criminal appeals from district and county courts.",
     "Darlene Byrne","D","true",
     "Incumbent Chief Justice of the 3rd Court of Appeals. Democrat defending her seat against Republican challenger Cory Liu in this Central Texas appellate race.",
     D+"1UAvHXWFdo6HbbLBjz--K-7PxFF1k8YuB"],
    ["3rd Court of Appeals — Chief Justice","judicial","","",
     "The 3rd Court of Appeals serves Central Texas including Hays County — hears civil and criminal appeals from district and county courts.",
     "Cory Liu","R","true",
     "Republican nominee for 3rd Court of Appeals Chief Justice. Challenging Democratic incumbent Darlene Byrne for this Central Texas appellate seat.",
     "https://texasjudges.org/wp-content/uploads/2025/12/Cory_Liu.jpeg"],

    ["428th District Judge","judicial","","",
     "Presides over the 428th Judicial District Court in Hays County — handles felony criminal cases and civil matters.",
     "TBD — Democratic Nominee","D","false","Democratic nominee to be confirmed.",""],
    ["428th District Judge","judicial","","",
     "Presides over the 428th Judicial District Court in Hays County — handles felony criminal cases and civil matters.",
     "TBD — Republican Nominee","R","false","Republican nominee to be confirmed.",""],

    ["County Court at Law No. 1","judicial","","",
     "Handles misdemeanor criminal cases, civil disputes, and probate matters in Hays County.",
     "TBD — Democratic Nominee","D","false","Democratic nominee to be confirmed.",""],
    ["County Court at Law No. 1","judicial","","",
     "Handles misdemeanor criminal cases, civil disputes, and probate matters in Hays County.",
     "TBD — Republican Nominee","R","false","Republican nominee to be confirmed.",""],

    ["County Court at Law No. 2","judicial","","",
     "Handles misdemeanor criminal cases, civil disputes, Juvenile Court, and Veterans Treatment Court.",
     "Chris Johnson","D","true",
     "Incumbent County Court at Law No. 2 judge who won the May 26 Democratic runoff. Running for another term against Republican nominee Charmaine Wilde.",
     D+"1-atwM5aVvoN-T_CrXJN0J7CFpFc5Vq-u"],
    ["County Court at Law No. 2","judicial","","",
     "Handles misdemeanor criminal cases, civil disputes, Juvenile Court, and Veterans Treatment Court.",
     "Charmaine Wilde","R","true",
     "Republican nominee for Hays County Court at Law No. 2. Challenging Democratic incumbent Chris Johnson in this county judicial race.",
     "https://smcorridornews.com/wp-content/uploads/2021/12/Screen-Shot-2021-12-27-at-1.08.25-PM-300x190.png"]

  ];

  var allRows = [headers];
  for (var i = 0; i < candidates.length; i++) {
    var c = candidates[i];
    var pos = (c[6] === "D") ? [A,A,A,A,A,A,A,A,A,A] : [B,B,B,B,B,B,B,B,B,B];
    var row = [c[0],c[1],c[2],c[3],c[5],c[6],c[7],c[4],c[8],"",c[9]||""];
    for (var k = 0; k < pos.length; k++) { row.push(pos[k]); }
    allRows.push(row);
  }

  stSheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);

  Logger.log("SUCCESS: wrote " + (allRows.length - 1) + " candidate rows with Drive photos for D candidates.");
  Logger.log("After running: File > Share > Publish to web to refresh the CSV.");
}
