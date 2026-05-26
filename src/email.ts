// Pulls fields off the current Outlook message and best-effort-parses them
// into Design PM project defaults.

import type { Priority } from "./types";

export type EmailContext = {
  subject: string;
  bodyText: string;
  senderName: string;
  senderEmail: string;
  receivedISO: string;
  detectedDueDate: string | null;
  detectedPriority: Priority | null;
  firstUrl: string | null;
};

export async function readCurrentEmail(): Promise<EmailContext> {
  const item = Office.context.mailbox.item;
  if (!item) throw new Error("No mail item in context.");

  const subject = item.subject ?? "";
  const from = item.from;
  const senderName = from?.displayName ?? "";
  const senderEmail = from?.emailAddress ?? "";
  const receivedISO = (item.dateTimeCreated ?? new Date()).toISOString();

  const bodyText = await new Promise<string>((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        resolve("");
      }
    });
  });

  return {
    subject,
    bodyText,
    senderName,
    senderEmail,
    receivedISO,
    detectedDueDate: detectDueDate(`${subject}\n${bodyText}`),
    detectedPriority: detectPriority(`${subject}\n${bodyText}`),
    firstUrl: detectFirstUrl(bodyText),
  };
}

function detectPriority(text: string): Priority | null {
  const t = text.toLowerCase();
  if (/\burgent\b|\basap\b|\beod\b|\brush\b/.test(t)) return "Urgent";
  if (/\bhigh priority\b|\bimportant\b|\bcritical\b/.test(t)) return "High";
  if (/\blow priority\b|\bwhenever\b|\bno rush\b/.test(t)) return "Low";
  return null;
}

function detectFirstUrl(text: string): string | null {
  const m = text.match(/https?:\/\/[^\s)>\]]+/);
  return m ? m[0] : null;
}

const MONTHS: Record<string, number> = {
  jan: 0, january: 0,
  feb: 1, february: 1,
  mar: 2, march: 2,
  apr: 3, april: 3,
  may: 4,
  jun: 5, june: 5,
  jul: 6, july: 6,
  aug: 7, august: 7,
  sep: 8, sept: 8, september: 8,
  oct: 9, october: 9,
  nov: 10, november: 10,
  dec: 11, december: 11,
};

function detectDueDate(text: string): string | null {
  const t = text.toLowerCase();
  const phrases = ["due by", "needed by", "deadline", "due date", "due on", "by eod"];
  let scan = t;
  for (const phrase of phrases) {
    const idx = scan.indexOf(phrase);
    if (idx >= 0) {
      const window = scan.slice(idx, idx + 80);
      const parsed = parseLooseDate(window);
      if (parsed) return parsed;
    }
  }
  // Generic last-resort scan over the whole text.
  return parseLooseDate(t);
}

function parseLooseDate(text: string): string | null {
  // ISO yyyy-mm-dd
  const iso = text.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;

  // d Month yyyy / Month d, yyyy
  const dmy = text.match(
    /(\d{1,2})(?:st|nd|rd|th)?\s+(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?(?:\s+(\d{2,4}))?/i
  );
  if (dmy) {
    const day = +dmy[1];
    const mo = MONTHS[dmy[2].toLowerCase()];
    const yr = dmy[3] ? normaliseYear(+dmy[3]) : new Date().getFullYear();
    if (!isNaN(day) && mo !== undefined) return toIsoDate(yr, mo, day);
  }
  const mdy = text.match(
    /(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{2,4}))?/i
  );
  if (mdy) {
    const day = +mdy[2];
    const mo = MONTHS[mdy[1].toLowerCase()];
    const yr = mdy[3] ? normaliseYear(+mdy[3]) : new Date().getFullYear();
    if (!isNaN(day) && mo !== undefined) return toIsoDate(yr, mo, day);
  }

  // dd/mm/yyyy or dd-mm-yyyy (Australian default; sender is at WiseTech)
  const num = text.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/);
  if (num) {
    const day = +num[1];
    const mo = +num[2] - 1;
    const yr = normaliseYear(+num[3]);
    if (day >= 1 && day <= 31 && mo >= 0 && mo <= 11) return toIsoDate(yr, mo, day);
  }
  return null;
}

function normaliseYear(y: number): number {
  if (y < 100) return 2000 + y;
  return y;
}

function toIsoDate(year: number, monthIndex: number, day: number): string {
  const m = String(monthIndex + 1).padStart(2, "0");
  const d = String(day).padStart(2, "0");
  return `${year}-${m}-${d}`;
}
