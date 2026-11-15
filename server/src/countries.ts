/**
 * Countries for SMS_ALLOWED_COUNTRIES: country calling codes (ITU E.164) and the many ways an admin may write a
 * country – ISO codes ("DE", "DEU"), calling codes ("+49", "0049", "49"), names in German and English with or without
 * umlauts ("Österreich", "Oesterreich", "Austria"), common short forms ("UK", "USA") and groups ("EU", "EWR", "DACH").
 *
 * A phone number tells only its calling code, not its country: countries that share a code (USA, Canada and the
 * Caribbean +1, Russia and Kazakhstan +7, the UK with Jersey, Guernsey and the Isle of Man +44, …) cannot be told
 * apart – allowing one of them allows the whole code.
 */

/** ISO 3166-1 alpha-2, alpha-3, calling code. The calling codes are prefix-free: no code starts with another one. */
const TABLE =
  "AD AND 376,AE ARE 971,AF AFG 93,AG ATG 1,AI AIA 1,AL ALB 355,AM ARM 374,AO AGO 244,AQ ATA 672,AR ARG 54," +
  "AS ASM 1,AT AUT 43,AU AUS 61,AW ABW 297,AX ALA 358,AZ AZE 994,BA BIH 387,BB BRB 1,BD BGD 880,BE BEL 32," +
  "BF BFA 226,BG BGR 359,BH BHR 973,BI BDI 257,BJ BEN 229,BL BLM 590,BM BMU 1,BN BRN 673,BO BOL 591,BQ BES 599," +
  "BR BRA 55,BS BHS 1,BT BTN 975,BW BWA 267,BY BLR 375,BZ BLZ 501,CA CAN 1,CC CCK 61,CD COD 243,CF CAF 236," +
  "CG COG 242,CH CHE 41,CI CIV 225,CK COK 682,CL CHL 56,CM CMR 237,CN CHN 86,CO COL 57,CR CRI 506,CU CUB 53," +
  "CV CPV 238,CW CUW 599,CX CXR 61,CY CYP 357,CZ CZE 420,DE DEU 49,DJ DJI 253,DK DNK 45,DM DMA 1,DO DOM 1," +
  "DZ DZA 213,EC ECU 593,EE EST 372,EG EGY 20,EH ESH 212,ER ERI 291,ES ESP 34,ET ETH 251,FI FIN 358,FJ FJI 679," +
  "FK FLK 500,FM FSM 691,FO FRO 298,FR FRA 33,GA GAB 241,GB GBR 44,GD GRD 1,GE GEO 995,GF GUF 594,GG GGY 44," +
  "GH GHA 233,GI GIB 350,GL GRL 299,GM GMB 220,GN GIN 224,GP GLP 590,GQ GNQ 240,GR GRC 30,GT GTM 502,GU GUM 1," +
  "GW GNB 245,GY GUY 592,HK HKG 852,HN HND 504,HR HRV 385,HT HTI 509,HU HUN 36,ID IDN 62,IE IRL 353,IL ISR 972," +
  "IM IMN 44,IN IND 91,IO IOT 246,IQ IRQ 964,IR IRN 98,IS ISL 354,IT ITA 39,JE JEY 44,JM JAM 1,JO JOR 962," +
  "JP JPN 81,KE KEN 254,KG KGZ 996,KH KHM 855,KI KIR 686,KM COM 269,KN KNA 1,KP PRK 850,KR KOR 82,KW KWT 965," +
  "KY CYM 1,KZ KAZ 7,LA LAO 856,LB LBN 961,LC LCA 1,LI LIE 423,LK LKA 94,LR LBR 231,LS LSO 266,LT LTU 370," +
  "LU LUX 352,LV LVA 371,LY LBY 218,MA MAR 212,MC MCO 377,MD MDA 373,ME MNE 382,MF MAF 590,MG MDG 261,MH MHL 692," +
  "MK MKD 389,ML MLI 223,MM MMR 95,MN MNG 976,MO MAC 853,MP MNP 1,MQ MTQ 596,MR MRT 222,MS MSR 1,MT MLT 356," +
  "MU MUS 230,MV MDV 960,MW MWI 265,MX MEX 52,MY MYS 60,MZ MOZ 258,NA NAM 264,NC NCL 687,NE NER 227,NF NFK 672," +
  "NG NGA 234,NI NIC 505,NL NLD 31,NO NOR 47,NP NPL 977,NR NRU 674,NU NIU 683,NZ NZL 64,OM OMN 968,PA PAN 507," +
  "PE PER 51,PF PYF 689,PG PNG 675,PH PHL 63,PK PAK 92,PL POL 48,PM SPM 508,PN PCN 64,PR PRI 1,PS PSE 970," +
  "PT PRT 351,PW PLW 680,PY PRY 595,QA QAT 974,RE REU 262,RO ROU 40,RS SRB 381,RU RUS 7,RW RWA 250,SA SAU 966," +
  "SB SLB 677,SC SYC 248,SD SDN 249,SE SWE 46,SG SGP 65,SH SHN 290,SI SVN 386,SJ SJM 47,SK SVK 421,SL SLE 232," +
  "SM SMR 378,SN SEN 221,SO SOM 252,SR SUR 597,SS SSD 211,ST STP 239,SV SLV 503,SX SXM 1,SY SYR 963,SZ SWZ 268," +
  "TC TCA 1,TD TCD 235,TG TGO 228,TH THA 66,TJ TJK 992,TK TKL 690,TL TLS 670,TM TKM 993,TN TUN 216,TO TON 676," +
  "TR TUR 90,TT TTO 1,TV TUV 688,TW TWN 886,TZ TZA 255,UA UKR 380,UG UGA 256,US USA 1,UY URY 598,UZ UZB 998," +
  "VA VAT 39,VC VCT 1,VE VEN 58,VG VGB 1,VI VIR 1,VN VNM 84,VU VUT 678,WF WLF 681,WS WSM 685,XK XKX 383," +
  "YE YEM 967,YT MYT 262,ZA ZAF 27,ZM ZMB 260,ZW ZWE 263";

const COUNTRIES = TABLE.split(",").map((row) => {
  const [region, alpha3, code] = row.split(" ");
  return { region, alpha3, code };
});
const CODE_OF = new Map(COUNTRIES.map((c) => [c.region, c.code]));
const CODES = new Set(COUNTRIES.map((c) => c.code));

/** The country a shared calling code stands for when the admin writes the code ("+1") instead of a country. */
const MAIN_REGION: Record<string, string> = {
  "1": "US", "7": "RU", "39": "IT", "44": "GB", "47": "NO", "61": "AU", "64": "NZ", "212": "MA", "262": "RE",
  "358": "FI", "590": "GP", "599": "CW", "672": "NF",
};

const EU = "AT BE BG CY CZ DE DK EE ES FI FR GR HR HU IE IT LT LU LV MT NL PL PT RO SE SI SK".split(" ");
const GROUPS: Record<string, string[]> = {
  eu: EU,
  europaischeunion: EU,
  europeanunion: EU,
  ewr: [...EU, "IS", "LI", "NO"],
  eea: [...EU, "IS", "LI", "NO"],
  dach: ["DE", "AT", "CH"],
};

/** Short and old forms the ISO names and Intl do not cover (keys as made by key()). */
const ALIASES: Record<string, string> = {
  uk: "GB", greatbritain: "GB", grossbritannien: "GB", england: "GB", scotland: "GB", schottland: "GB", wales: "GB",
  northernireland: "GB", nordirland: "GB", usa: "US", unitedstatesofamerica: "US", vereinigtestaatenvonamerika: "US",
  amerika: "US", america: "US", brd: "DE", bundesrepublikdeutschland: "DE", holland: "NL", czechrepublic: "CZ",
  tschechischerepublik: "CZ", turkey: "TR", turkei: "TR", russianfederation: "RU", southkorea: "KR", korea: "KR",
  northkorea: "KP", macedonia: "MK", mazedonien: "MK", swaziland: "SZ", burma: "MM", birma: "MM", ivorycoast: "CI",
  elfenbeinkuste: "CI", vatican: "VA", vatikan: "VA", holysee: "VA", capeverde: "CV", kapverden: "CV",
};

/** Lower case, no accents or umlaut dots, ß → ss, only letters and digits ("Côte d’Ivoire" → "cotedivoire"). */
const key = (value: string) =>
  value.normalize("NFD").replace(/\p{M}/gu, "").replace(/ß/g, "ss").toLowerCase().replace(/[^a-z0-9]/g, "");
/** German spelling without umlauts: "Österreich" → "Oesterreich". */
const UMLAUTS: Record<string, string> = { ä: "ae", Ä: "Ae", ö: "oe", Ö: "Oe", ü: "ue", Ü: "Ue" };
const transliterate = (value: string) => value.replace(/[äÄöÖüÜ]/g, (c) => UMLAUTS[c] ?? c);

/** Name of a country in a language ("DE", "de" → "Deutschland"); the ISO code if Intl does not know it. */
export function countryName(region: string, language: string): string {
  try {
    return new Intl.DisplayNames([language, "en"], { type: "region" }).of(region) ?? region;
  } catch {
    return region;
  }
}

let names: Map<string, string> | undefined;
/** Key → ISO code for every name in German, English and the extra language, with and without umlauts. */
function nameIndex(extraLanguage: string): Map<string, string> {
  if (names) return names;
  names = new Map();
  const languages = [...new Set(["de", "en", extraLanguage])];
  for (const { region, alpha3 } of COUNTRIES) {
    names.set(region.toLowerCase(), region);
    names.set(alpha3.toLowerCase(), region);
    for (const language of languages) {
      let display: Intl.DisplayNames;
      try {
        display = new Intl.DisplayNames([language], { type: "region", fallback: "none" });
      } catch {
        continue; // a language Intl does not know
      }
      const name = display.of(region);
      if (!name) continue;
      for (const spelling of [name, transliterate(name)]) if (!names.has(key(spelling))) names.set(key(spelling), region);
    }
  }
  for (const [alias, region] of Object.entries(ALIASES)) names.set(alias, region);
  return names;
}

/**
 * The countries (ISO codes) one entry of SMS_ALLOWED_COUNTRIES stands for, or null if it names none.
 * "+49" / "0049" / "49" → ["DE"], "Österreich" / "AUT" / "at" → ["AT"], "EU" → the 27 member states.
 */
export function resolveCountry(entry: string, language = "de"): string[] | null {
  const text = entry.trim();
  const code = /^(?:\+|00)?\s*(\d{1,3})$/.exec(text)?.[1];
  if (code) return CODES.has(code) ? [MAIN_REGION[code] ?? COUNTRIES.find((c) => c.code === code)!.region] : null;
  const k = key(text);
  if (GROUPS[k]) return GROUPS[k];
  const region = nameIndex(language).get(k);
  return region ? [region] : null;
}

/**
 * SMS_ALLOWED_COUNTRIES → ISO codes, in the order written and without duplicates; [] = all countries. Entries are
 * separated by commas, semicolons or line breaks (names may contain spaces: "Vereinigtes Königreich"). Throws on an
 * entry that names no country, so a typo does not silently block or allow numbers.
 */
export function parseCountries(raw: string, language = "de"): string[] {
  const regions: string[] = [];
  for (const entry of raw.split(/[,;\n]/).map((e) => e.trim()).filter(Boolean)) {
    const found = resolveCountry(entry, language);
    if (!found) {
      throw new Error(
        `Environment variable SMS_ALLOWED_COUNTRIES: unknown country "${entry}" – use ISO codes (DE, AT), calling codes (+49) or names (Deutschland, Austria)`,
      );
    }
    for (const region of found) if (!regions.includes(region)) regions.push(region);
  }
  return regions;
}

/** Calling code of a country ("DE" → "49"). */
export const callingCodeOf = (region: string): string | undefined => CODE_OF.get(region);

/** Calling code of a number in E.164 ("+491701234567" → "49"), or null if no country has it. */
export function callingCode(e164: string): string | null {
  const digits = e164.replace(/^\+/, "");
  for (const length of [1, 2, 3]) {
    if (CODES.has(digits.slice(0, length))) return digits.slice(0, length);
  }
  return null;
}
