// Shared by floor-plan.html and seat-finder.html.
// Reads the same saved plan as seating.html (localStorage "wedding-seating-v1")
// and keeps the table → floor-spot mapping in "wedding-floorplan-v1".
const FP = (() => {
  const SEAT_STORE = "wedding-seating-v1", FP_STORE = "wedding-floorplan-v1", MAX_SEATS = 12;

  // Copy of the base sheet in seating.html — only used when this device has no saved seating plan yet.
  const SEED = [{"name": "Ejean's Family 6", "guests": ["Er gu", "Guzhang", "Chee kit", "Chee Kit's wife", "Chee Kit's Daughter x2", "Chee Hou", "Chee Hou's wife", "Joo Peng", "Chee Kang"], "seats": 10}, {"name": "Ejean's Family 5", "guests": ["Grandpa", "Grandma", "Yiyi", "Yizhang", "Yun Wei", "Kent If", "Cuebee", "Cue Bee's Son x2", "Ying ying", "Popo's sister (Unconfirmed)"], "seats": 11}, {"name": "Ejean's Family 4", "guests": ["Da jiu", "Yong yong", "Jeff's Mom", "Sam suk po", "Ah yong", "Ah yong's wife", "Ah Yong's Son", "Fong Yi", "Fong Yi's Husband", "Fong Yi's Son"], "seats": 10}, {"name": "Ejean's Family 3", "guests": ["Winnie", "Winnie's Husband", "Winnie's 4x children", "Li Shyun", "Li Shyun's 2x children"], "seats": 10}, {"name": "Ejean's Family 2", "guests": ["Fifth grandaunt", "Fifth grandaunt's husband", "Sixth grandaunt", "Sixth grandaunt's husband", "Sixth grandaunt's Son", "Sixth grandaunt's Daughter In Law", "Sixth grandaunt's Son", "Seventh grandaunt", "Seventh grandaunt' husband", "Jun", "Jun's wife", "Jun's child"], "seats": 12}, {"name": "Ejean's Family 1", "guests": ["Xiao gu", "Xiao Gu's Youngest daughter", "Xiao Gu Youngest daughter's husband", "Xiao Gu's Youngest daughter baby", "Da gu", "Da gu husband", "Ba mei", "Suk Mei"], "seats": 10}, {"name": "Mom's friends", "guests": ["Mom's friends x10"], "seats": 10}, {"name": "Mama & Dedey", "guests": ["Mama friends x8", "Mama", "Dedey", "Kelly"], "seats": 11}, {"name": "Friends", "guests": ["Catherine", "You Soon", "Sharmaine", "Wei Xiang", "Chloe", "Chloe's Partner", "Man Teng", "Jing Xuan", "Wen Di", "Pui San"], "seats": 10}, {"name": "Colleagues", "guests": ["Agnes", "Zhi Ying", "Zhao tian", "Sengja", "Hsao Kham", "Htay Htay", "Evelyne", "Stanley", "Brandon"], "seats": 10}, {"name": "Ballet table", "guests": ["Ying Yi", "Pui Yee", "Pui Yee's Husband", "Teacher", "Ballet girls 8x"], "seats": 12}, {"name": "JL's Family 1", "guests": ["Bheng Lim", "Meini", "Bheng lim's 4x children", "Bheng Siang", {"name": "Bheng Siang's children", "pax": 2}, "Bheng liang", "Jun", "Chiu Yan"], "seats": 12}, {"name": "VIP Table", "guests": ["Jing Long", "Ejean", "Jing Long's Mother", "Jing Long's Father", "Ejean's Mom", "Jeffrey (Didi)", "Ejean's 婆婆", "Jing long's 姑姑", "Shiow Han", "Guo Xuan"], "seats": 10}, {"name": "JL Family 2", "guests": ["Shiow Han's Mom", "Shiow Han's Dad", "Aunt Jacqueline (四姨）", "Aunt Carol (三姨）", "Uncle Wei Yong (舅舅)", "Aunt Audrey (大姨）", "Jing Xian", "Jing Kai", "Jing Wei", "Shiow Han's Brother"], "seats": 10}];
  const SEED_SIG = JSON.stringify(SEED).length + ":" + [...JSON.stringify(SEED)].reduce((h,c) => (h*31 + c.charCodeAt(0)) >>> 0, 7);

  // Table spots, traced from the venue drawing dated 22 Nov 2026 (the drawing labels the last one 15; it is table 14).
  const SPOTS = [
    {id:"VIP", x:198, y:303, big:true},
    {id:"1", x:400, y:62}, {id:"2", x:322, y:58}, {id:"3", x:233, y:58}, {id:"5", x:150, y:100},
    {id:"6", x:233, y:138}, {id:"7", x:320, y:138}, {id:"8", x:400, y:135},
    {id:"9", x:403, y:465}, {id:"10", x:322, y:472}, {id:"11", x:230, y:490},
    {id:"12", x:229, y:575}, {id:"13", x:325, y:560}, {id:"14", x:407, y:543},
  ];
  // Default placement: spot → table name (as named in the seating plan).
  // To update it, export the floor plan to Excel and use its "Floor plan" sheet.
  // From "Wedding - Floor plan 8 nov.xlsx".
  const DEFAULT_MAP = {"VIP": "VIP Table", "1": "Ballet table", "2": "Mom's friends", "3": "Mama & Dedey", "5": "Colleagues", "6": "JL's Family 1", "7": "JL Family 2", "8": "Friends", "9": "Ejean's Family 4", "10": "Ejean's Family 5", "11": "Ejean's Family 6", "12": "Ejean's Family 2", "13": "Ejean's Family 3", "14": "Ejean's Family 1"};
  const DEFAULT_SIG = JSON.stringify(DEFAULT_MAP);
  const spotLabel = id => id === "VIP" ? "VIP table" : "Table " + id;

  const normName = s => String(s).toLowerCase().replace(/['’\s]/g, "");
  const esc = s => String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));

  function parsePax(name){
    const s = name.toLowerCase(); let m;
    if ((m = s.match(/(\d+)\s*[-–~]\s*(\d+)\s*x/)) || (m = s.match(/x\s*(\d+)\s*[-–~]\s*(\d+)/))) return Math.max(+m[1], +m[2]);
    if ((m = s.match(/x\s*(\d+)/)) || (m = s.match(/(\d+)\s*x\b/)) || (m = s.match(/(\d+)\s*x/))) return Math.max(1, +m[1]);
    if (s.includes("+")) return s.split("+").filter(p => p.trim()).length;
    if (/\b(parents|couple)\b/.test(s)) return 2;
    return 1;
  }

  let uid = 1;
  const nid = p => p + (uid++).toString(36) + Math.random().toString(36).slice(2,5);
  function fromSeed(){
    const tables = [], guests = [];
    SEED.forEach(c => {
      const t = { id: nid("t"), name: c.name, seats: 10 }; let pax = 0;
      c.guests.forEach(g => {   // a guest is a name, or {name, pax} when the headcount isn't in the name
      const name = typeof g === "string" ? g : g.name, p = typeof g === "string" ? parsePax(g) : g.pax;
      pax += p; guests.push({ id: nid("g"), name, pax: p, table: t.id });
    });
      t.seats = Math.min(MAX_SEATS, Math.max(1, c.seats || Math.max(10, pax)));
      tables.push(t);
    });
    return { tables, guests, seedSig: SEED_SIG };
  }

  // persist=true: on a fresh device, save the base sheet so seating.html and this page share table ids.
  function loadSeating(persist = true){
    let s = null;
    try { s = JSON.parse(localStorage.getItem(SEAT_STORE)); } catch(e) {}
    if (!s || !s.tables || s.seedSig !== SEED_SIG) {   // same rule as seating.html: a newer base sheet replaces older saved progress
      s = fromSeed();
      if (persist) try { localStorage.setItem(SEAT_STORE, JSON.stringify(s)); } catch(e) {}
    }
    return s;
  }

  // Returns { spotId: tableId }. Stored entries carry the table name too, so a reset/import
  // in seating.html (which re-creates table ids) keeps the placement.
  function loadMap(seating){
    let raw = null;
    try { raw = JSON.parse(localStorage.getItem(FP_STORE)); } catch(e) {}
    // A newer default baked into this file replaces placements saved against the old one
    // (placements saved before defaults existed carry no stamp and are kept)
    if (!raw || !raw.spots || (raw.defaultSig !== undefined && raw.defaultSig !== DEFAULT_SIG)) raw = defaultRaw();
    const map = {}, used = new Set();
    Object.entries(raw.spots).forEach(([sid, ref]) => {
      if (sid === "15") sid = "14";   // spot was called 15 before 8 Nov
      if (!SPOTS.some(s => s.id === sid) || !ref) return;
      const t = seating.tables.find(t => t.id === ref.id) || seating.tables.find(t => normName(t.name) === normName(ref.name));
      if (t && !used.has(t.id)) { map[sid] = t.id; used.add(t.id); }
    });
    return map;
  }
  function saveMap(map, seating){
    const spots = {};
    Object.entries(map).forEach(([sid, tid]) => {
      const t = seating.tables.find(t => t.id === tid);
      if (t) spots[sid] = { id: t.id, name: t.name };
    });
    try { localStorage.setItem(FP_STORE, JSON.stringify({ spots, defaultSig: DEFAULT_SIG })); } catch(e) {}
  }
  const defaultRaw = () => ({ spots: Object.fromEntries(Object.entries(DEFAULT_MAP).map(([sid, name]) => [sid, { name }])) });
  function defaultMap(seating){
    const map = {};
    Object.entries(DEFAULT_MAP).forEach(([sid, name]) => {
      const t = seating.tables.find(t => normName(t.name) === normName(name));
      if (t) map[sid] = t.id;
    });
    return map;
  }

  // Everything a page needs to show, in floor order: placed tables first, then the rest.
  function directory(seating, map){
    const spotOf = tid => Object.keys(map).find(k => map[k] === tid) || null;
    const order = id => id ? SPOTS.findIndex(s => s.id === id) : 99;
    return seating.tables.map(t => {
      const gs = seating.guests.filter(g => g.table === t.id);
      const sid = spotOf(t.id);
      return { tid: t.id, sid, name: t.name, seats: t.seats, pax: gs.reduce((a,g) => a + g.pax, 0),
               guests: gs.map(g => ({ name: g.name, pax: g.pax })) };
    }).sort((a,b) => order(a.sid) - order(b.sid) || (+a.sid || 0) - (+b.sid || 0));
  }

  // ---------- shareable snapshot (#z=… in the seat-finder URL) ----------
  const b64u = bytes => { let s = ""; bytes.forEach(b => s += String.fromCharCode(b)); return btoa(s).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,""); };
  const unb64u = str => Uint8Array.from(atob(str.replace(/-/g,"+").replace(/_/g,"/")), c => c.charCodeAt(0));
  async function encode(dir){
    const json = JSON.stringify({ v:1, t: dir.map(d => [d.sid || "", d.name, d.seats, d.guests.map(g => g.pax === parsePax(g.name) ? g.name : [g.name, g.pax])]) });
    if (window.CompressionStream) {
      const stream = new Blob([json]).stream().pipeThrough(new CompressionStream("deflate-raw"));
      return "z=" + b64u(new Uint8Array(await new Response(stream).arrayBuffer()));
    }
    return "j=" + b64u(new TextEncoder().encode(json));
  }
  async function decode(hash){
    const m = String(hash).match(/^#?([zj])=([\w-]+)/); if (!m) return null;
    let json;
    if (m[1] === "z") {
      const stream = new Blob([unb64u(m[2])]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      json = await new Response(stream).text();
    } else json = new TextDecoder().decode(unb64u(m[2]));
    const data = JSON.parse(json);
    return data.t.map(([sid, name, seats, names], i) => {
      const guests = names.map(n => Array.isArray(n) ? { name: n[0], pax: n[1] } : { name: n, pax: parsePax(n) });
      return { tid: "s" + i, sid: sid === "15" ? "14" : sid || null, name, seats, pax: guests.reduce((a,g) => a + g.pax, 0), guests };
    });
  }

  // ---------- drawing ----------
  const treads = (x1, x2, y1, y2, step) => { let d = ""; for (let y = y1; y <= y2; y += step) d += `M${x1} ${y}H${x2}`; return `<path class="fine" d="${d}"/>`; };

  function roomSVG(){
    return `
    <defs>
      <pattern id="fpGrid" width="20" height="20" patternUnits="userSpaceOnUse"><path d="M20 0H0V20" class="gridline"/></pattern>
    </defs>
    <rect x="0" y="0" width="1180" height="660" fill="url(#fpGrid)"/>
    <!-- outer walls, drawn as a double line -->
    <g>
      <path class="wall" d="M15 12H1180M15 12V170L70 560Q82 640 140 636L1125 420H1140V212H1180"/>
      <path class="wall-in" d="M15 12H1180M15 12V170L70 560Q82 640 140 636L1125 420H1140V212H1180"/>
    </g>
    <!-- left stair core -->
    <rect class="thin" x="20" y="16" width="62" height="140" rx="31"/>
    <path class="fine" d="M51 34V138"/>${treads(32, 70, 40, 132, 8)}
    <!-- left service block -->
    <rect class="thin" x="92" y="205" width="58" height="200"/>
    <path class="thick" d="M150 214V396"/>
    <path class="fine" d="M24 236H92M24 236V376M24 376H70"/>
    <path class="fine" d="M70 376A22 22 0 0 1 92 398"/>
    <!-- central stair -->
    <rect class="thin" x="455" y="42" width="85" height="122"/>
    <path class="fine" d="M497 46V160"/>${treads(459, 536, 54, 156, 9)}
    <text class="ann" x="476" y="38" text-anchor="middle">dn</text><text class="ann" x="518" y="38" text-anchor="middle">up</text>
    <rect class="dash" x="430" y="44" width="25" height="238"/>
    <path class="thin" d="M468 212H540V170M540 212H610"/>
    <!-- lifts -->
    <rect class="thick" x="612" y="10" width="68" height="66"/>
    <rect class="thick" x="612" y="76" width="68" height="68"/>
    <rect class="thick" x="612" y="144" width="68" height="68"/>
    <text class="ann" x="646" y="47" text-anchor="middle">Bomba</text>
    <path class="fine" d="M620 84L672 136M672 84L620 136M620 152L672 204M672 152L620 204"/>
    <!-- grand stair -->
    <path class="thin" d="M690 30H862M700 72Q781 28 858 72"/>
    <g transform="rotate(-9 726 132)"><rect class="thin" x="700" y="86" width="52" height="94"/>${treads(700, 752, 92, 176, 7)}</g>
    <g transform="rotate(9 824 132)"><rect class="thin" x="796" y="86" width="56" height="94"/>${treads(796, 852, 92, 176, 7)}</g>
    <path class="thin" d="M862 212H1140"/>
    <!-- corridor -->
    <path class="thin" d="M480 262H1135M480 250H595M623 250H658M686 250H878M906 250H942M970 250H1135"/>
    <rect class="colo" x="463" y="257" width="32" height="22"/>
    <rect class="colo" x="595" y="246" width="26" height="26"/><rect class="colo" x="658" y="246" width="26" height="26"/>
    <rect class="colo" x="878" y="246" width="26" height="26"/><rect class="colo" x="942" y="246" width="26" height="26"/>
    <path class="fine" d="M1128 272V408M1132 290H1140M1132 310H1140M1132 330H1140M1132 350H1140M1132 370H1140"/>
    <!-- columns -->
    <g class="col">
      <rect x="88" y="14" width="10" height="24"/><rect x="272" y="14" width="10" height="30"/><rect x="462" y="14" width="10" height="28"/>
      <rect x="862" y="14" width="10" height="30"/><rect x="1052" y="14" width="10" height="24"/>
      <rect x="82" y="170" width="14" height="44"/><rect x="462" y="170" width="12" height="44"/>
      <rect x="862" y="170" width="12" height="42"/><rect x="1050" y="170" width="12" height="42"/>
      <rect x="80" y="383" width="16" height="50"/>
      <rect x="88" y="560" width="8" height="22"/><rect x="272" y="548" width="10" height="30"/><rect x="462" y="556" width="8" height="24"/>
      <rect x="676" y="497" width="8" height="22"/><rect x="862" y="400" width="8" height="22"/><rect x="1053" y="398" width="8" height="22"/>
    </g>
    <rect class="colo" x="456" y="378" width="22" height="58"/><rect class="col" x="461" y="384" width="12" height="46"/>
    <rect class="colo" x="668" y="380" width="22" height="50"/><rect class="col" x="673" y="386" width="12" height="38"/>
    <!-- white carpet -->
    <rect class="carpet" x="245" y="283" width="415" height="65"/>
    <text class="carpet-t" x="452" y="321" text-anchor="middle">White carpet</text>
    <!-- title block -->
    <text class="title-date" x="1110" y="606" text-anchor="end">22 Nov 2026</text>
    <text class="ann" x="1110" y="626" text-anchor="end">Banquet floor, Jing Long &amp; Ejean</text>`;
  }

  function vendorsSVG(){
    return `<g class="vendor">
      <rect class="v-shape" x="760" y="186" width="100" height="30" rx="15"/>
      <text class="v-text" x="810" y="205" text-anchor="middle">Registration</text>
      <path class="chev" d="M760 290L812 358L760 426"/>
      <text class="v-text" transform="translate(767 358) rotate(90)" text-anchor="middle">Photo booth</text>
      <text class="v-text" transform="translate(836 358) rotate(90)" text-anchor="middle">Gallery area</text>
      <rect class="v-solid" x="1110" y="280" width="8" height="76"/>
      <text class="v-text" transform="translate(1096 318) rotate(90)" text-anchor="middle">Tea ceremony backdrop</text>
      <rect class="v-shape" x="885" y="365" width="125" height="40" rx="6"/>
      <text class="v-text" x="947" y="389" text-anchor="middle">Pastries</text>
      <ellipse class="v-shape" cx="800" cy="464" rx="46" ry="18"/>
      <text class="v-text" x="800" y="468" text-anchor="middle">Gelato</text>
      <rect class="v-shape" x="545" y="452" width="125" height="40" rx="6"/>
      <text class="v-text" x="607" y="476" text-anchor="middle">Perfume station</text>
    </g>`;
  }

  // One table spot. info = directory entry or null for an empty spot.
  function spotSVG(s, info, cls = ""){
    const big = !!s.big, R = big ? 38 : 28, r = big ? 27 : 19, cw = big ? 11 : 9.5, chh = 7;
    const seats = info ? info.seats : 10, pax = info ? info.pax : 0, total = Math.max(seats, pax);
    let chairs = "";
    for (let i = 0; i < total; i++){
      const a = -Math.PI/2 + i * 2 * Math.PI / total;
      const kind = i >= seats ? "over" : i < pax ? "on" : "";
      chairs += `<rect class="chair ${kind}" x="${-cw/2}" y="${-chh/2}" width="${cw}" height="${chh}" rx="2.5"
        transform="translate(${(R*Math.cos(a)).toFixed(1)} ${(R*Math.sin(a)).toFixed(1)}) rotate(${(a*180/Math.PI+90).toFixed(1)})"/>`;
    }
    const label = s.id, state = info ? "filled" : "empty", over = info && pax > seats ? " over" : "";
    const aria = info ? `${spotLabel(s.id)}: ${info.name}, ${pax} of ${seats} seats` : `${spotLabel(s.id)}: empty`;
    return `<g class="spot ${state}${over} ${cls}" data-spot="${s.id}" transform="translate(${s.x} ${s.y})" tabindex="0" role="button" aria-label="${esc(aria)}">
      <circle class="hit" r="${R + 8}"/>
      <circle class="halo" r="${R + 9}"/>
      ${chairs}
      <circle class="disc" r="${r}"/>
      <text class="num${big ? " vip" : ""}" y="1">${label}</text>
    </g>`;
  }

  return { SPOTS, SEAT_STORE, FP_STORE, spotLabel, normName, esc, parsePax, loadSeating, loadMap, saveMap, defaultMap, directory,
           encode, decode, roomSVG, vendorsSVG, spotSVG };
})();
