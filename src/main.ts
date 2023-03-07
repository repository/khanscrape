import { endOfDay, format, intervalToDuration, isBefore, parseISO } from "date-fns";
import { utcToZonedTime } from "date-fns-tz";
import XLSX from "exceljs";
import fs from "fs/promises";
import path from "path";
import { executablePath } from "puppeteer";
import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import "puppeteer-extra-plugin-stealth/evasions/chrome.app";
import "puppeteer-extra-plugin-stealth/evasions/chrome.csi";
import "puppeteer-extra-plugin-stealth/evasions/chrome.loadTimes";
import "puppeteer-extra-plugin-stealth/evasions/chrome.runtime";
import "puppeteer-extra-plugin-stealth/evasions/defaultArgs";
import "puppeteer-extra-plugin-stealth/evasions/iframe.contentWindow";
import "puppeteer-extra-plugin-stealth/evasions/media.codecs";
import "puppeteer-extra-plugin-stealth/evasions/navigator.hardwareConcurrency";
import "puppeteer-extra-plugin-stealth/evasions/navigator.languages";
import "puppeteer-extra-plugin-stealth/evasions/navigator.permissions";
import "puppeteer-extra-plugin-stealth/evasions/navigator.plugins";
import "puppeteer-extra-plugin-stealth/evasions/navigator.vendor";
import "puppeteer-extra-plugin-stealth/evasions/navigator.webdriver";
import "puppeteer-extra-plugin-stealth/evasions/sourceurl";
import "puppeteer-extra-plugin-stealth/evasions/user-agent-override";
import "puppeteer-extra-plugin-stealth/evasions/webgl.vendor";
import "puppeteer-extra-plugin-stealth/evasions/window.outerdimensions";
import "puppeteer-extra-plugin-user-data-dir";
import "puppeteer-extra-plugin-user-preferences";
import { ClassSubjectMasteryProgress } from "./interfaces/ClassSubjectMasteryProgress";
import { GetClassList } from "./interfaces/GetClassList";

puppeteer.use(StealthPlugin());

declare global {
  interface Window {
    setOverlay: (message: string) => void;
  }
}

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

function toLetter(index: number) {
  return String.fromCharCode(64 + index);
}

async function main() {
  const cwd = process.cwd();

  const outDir = path.resolve(cwd, "out");
  const configDir = path.resolve(cwd, "config");

  await Promise.all([outDir, configDir].map((dir) => fs.mkdir(dir, { recursive: true })));

  const cookiesPath = path.resolve(configDir, "cookies.json");
  const selectedClassesPath = path.resolve(configDir, "selectedClasses.json");

  let cookies: any;
  try {
    cookies = JSON.parse(await fs.readFile(cookiesPath, "utf-8"));
  } catch (error) {
    console.log("No cookies found");
  }

  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    executablePath: executablePath(),
    args: ["--start-maximized"],
  });

  const page = (await browser.pages())[0] ?? (await browser.newPage());

  await page.evaluateOnNewDocument(async () => {
    document.addEventListener("DOMContentLoaded", async () => {
      const style = document.createElement("style");
      style.innerHTML = `
.__overlay {
  position: fixed;
  top: 25px;
  left: 25px;
  font-size: 2rem;
  z-index: 9999;
  background: rgba(72, 52, 212, 0.75);
  padding: 1rem;
  pointer-events: none;
  color: white;
}
`;
      document.head.appendChild(style);

      window.setOverlay = (message: string) => {
        const overlay = document.getElementById("overlayer") ?? document.createElement("div");
        overlay.id = "overlayer";

        if (message.length <= 0) {
          overlay.remove();
          return;
        }

        overlay.classList.add("__overlay");
        overlay.innerText = message;
        document.body.appendChild(overlay);
      };
    });
  });

  if (cookies) {
    await page.setCookie(...cookies);
  }

  await page.goto("https://www.khanacademy.org/teacher/dashboard", { waitUntil: "domcontentloaded" });

  if (new URL(page.url()).pathname === "/login") {
    await page.evaluate(() => {
      window.setOverlay("Please login to Khan Academy");
    });
  }

  const classes = await page
    .waitForResponse((res) => new URL(res.url()).pathname === "/api/internal/graphql/getClassList", { timeout: 0 })
    .then((res) => res.json() as Promise<GetClassList>)
    .then(
      ({
        data: {
          coach: { studentLists },
        },
      }) =>
        studentLists
          .map(({ signupCode, name, countStudents, topics }) => ({
            id: signupCode,
            name,
            countStudents,
            topic: topics.map(({ title }) => title).join(", "),
          }))
          .sort((a, b) => (a.topic === b.topic ? a.name.localeCompare(b.name) : a.topic.localeCompare(b.topic))),
    );

  await fs.writeFile(cookiesPath, JSON.stringify(await page.cookies()));

  const classNameFormat = ({ name, topic, countStudents }: (typeof classes)[number]) =>
    `${name} (${topic}) - ${countStudents} students`;

  const formattedClasses = classes.map((c) => ({ ...c, formattedName: classNameFormat(c) }));

  const preslectedClassIds = await fs
    .readFile(selectedClassesPath, "utf-8")
    .then((data) => JSON.parse(data) as string[])
    .catch(() => [] as string[]);

  const selectorPage = await browser.newPage();

  const inputDateFormat = (date: Date) => format(date, "yyyy-MM-dd'T'HH:mm");
  const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const localNow = utcToZonedTime(new Date(), timeZone);

  await selectorPage.setContent(
    `
    <!DOCTYPE html>
    <html>
      <head>
        <title>Select Classes</title>
      </head>
      <body>
        <h1>Select Classes</h1>
        <form
        >
          ${formattedClasses
            .map(
              ({ id, formattedName }) => `
            <div>
              <input type="checkbox" data-class="true" id="${id}" name="${id}" value="${id}" ${
                preslectedClassIds.includes(id) ? "checked" : ""
              }>
              <label for="${id}">${formattedName}</label>
            </div>
`,
            )
            .join("")}
            <hr />
            <div>
              <input type="checkbox" id="schedule" name="schedule" value="schedule">
              <label for="schedule">Schedule? (Uncheck to run immediately)</label>
            </div>
            <div>
              <label for="scheduleTime">Time if scheduled</label>
              <input type="datetime-local" id="scheduleTime" name="scheduleTime"
              min="${inputDateFormat(localNow)}"
              value="${inputDateFormat(endOfDay(localNow))}">
            </div>
            <input type="submit" value="Submit">
          </form>
        </body>
      </html>`,
  );

  await selectorPage.waitForSelector("input[type=submit]", { timeout: 0 });

  const { selectedClassIds, scheduledAt } = await selectorPage.evaluate(() => {
    const form = document.querySelector("form") as HTMLFormElement;
    const checkboxesInput = Array.from(
      form.querySelectorAll("input[type=checkbox][data-class=true]"),
    ) as HTMLInputElement[];
    const scheduleInput = form.querySelector<HTMLInputElement>("input[name=schedule]");
    const scheduleTimeInput = form.querySelector<HTMLInputElement>("input[name=scheduleTime]");

    return new Promise<{
      selectedClassIds: string[];
      scheduledAt: string | null;
    }>((resolve) => {
      form.addEventListener("submit", (e) => {
        e.preventDefault();

        const selectedClassIds = checkboxesInput.filter((c) => c.checked).map((c) => c.value);
        const schedule = scheduleInput?.checked ?? false;
        const scheduleTime = scheduleTimeInput?.value ?? null;

        resolve({
          selectedClassIds: selectedClassIds,
          scheduledAt: schedule ? scheduleTime : null,
        });
      });
    });
  });

  await fs.writeFile(selectedClassesPath, JSON.stringify(selectedClassIds));

  await selectorPage.close();

  if (selectedClassIds.length <= 0) {
    let countdown = 5;

    while (countdown > 0) {
      await page.evaluate((countdown) => {
        window.setOverlay(`No classes selected. Closing in ${countdown}...`);
      }, countdown);

      await sleep(1000);
      countdown--;
    }

    return browser.close();
  }

  if (scheduledAt) {
    const parsed = parseISO(scheduledAt);

    while (isBefore(new Date(), parsed)) {
      const duration = intervalToDuration({ start: new Date(), end: parsed });
      const formatted = `${duration.hours?.toString().padStart(2, "0")}:${duration.minutes
        ?.toString()
        .padStart(2, "0")}:${duration.seconds?.toString().padStart(2, "0")}`;
      await page.evaluate((time) => {
        window.setOverlay(`Starting in ${time}... Please don't touch anything in the meantime!`);
      }, formatted);

      await sleep(1000);
    }
  }

  const classesWithMasteriesUnresolved = selectedClassIds.map((classId) => async () => {
    const selectedClass = formattedClasses.find((c) => c.id === classId);

    if (!selectedClass) {
      throw new Error(`Class with id ${classId} not found`);
    }

    await page.evaluate((name) => {
      window.setOverlay(`Scraping ${name}`);
    }, `"${selectedClass.name}" (${selectedClass.topic})`);

    await page.waitForSelector(`a[href='/teacher/class/${classId}']`, { timeout: 0 });
    await page.click(`a[href='/teacher/class/${classId}']`);

    if (selectedClass.countStudents <= 0) {
      return {
        ...selectedClass,
        masteries: [],
      };
    }

    await page.waitForSelector("a[data-test-id='nav-course-mastery-progress']", { timeout: 0 });
    await page.click("a[data-test-id='nav-course-mastery-progress']");

    const masteries = await page
      .waitForResponse(
        (res) =>
          new URL(res.url()).pathname === "/api/internal/graphql/ClassSubjectMasteryProgress" && res.status() === 200,
        { timeout: 0 },
      )
      .then((res) => res.json() as Promise<ClassSubjectMasteryProgress>)
      .then(
        ({
          data: {
            classroom: { students },
          },
        }) =>
          students.map(({ coachNickname, kaid, subjectProgress: { currentMastery: ocm, unitProgresses: ups } }) => ({
            id: kaid,
            name: coachNickname,
            overall: {
              percentage: ocm.pointsEarned / ocm.pointsAvailable,
              pointsAvailable: ocm.pointsAvailable,
              pointsEarned: ocm.pointsEarned,
            },
            units: ups.map(({ topic: { id }, currentMastery: upcm }) => ({
              id,
              percentage: upcm.pointsEarned / upcm.pointsAvailable,
              pointsAvailable: upcm.pointsAvailable,
              pointsEarned: upcm.pointsEarned,
            })),
          })),
      );

    await page.evaluate((students: number) => {
      window.setOverlay(`OK: Got data for ${students} student(s).`);
    }, masteries.length);

    return {
      ...selectedClass,
      masteries,
    };
  });

  const classesWithMasteries: Awaited<ReturnType<(typeof classesWithMasteriesUnresolved)[number]>>[] = [];

  for (const getClassWithMasteries of classesWithMasteriesUnresolved) {
    classesWithMasteries.push(await getClassWithMasteries());
    await page.click("[data-test-id='classroom-breadcrumb-link'] a");
  }

  await page.evaluate(() => {
    window.setOverlay("Saving file...");
  });

  const topicOrder = [
    "xb0caa40e7a222a7c",
    "x5432c6f7037ffefc",
    "xb1aee696f6974830",
    "x90c04f514dafe6d8",
    "xd8ffd65fe7b9701f",
    "x3a79bcb1ef658b56",
    "xf7faa64f661a9e62",
    "xb80ad6dc16530012",
  ];

  const workbook = new XLSX.Workbook();

  for (const classWithMastery of classesWithMasteries) {
    const worksheet = workbook.addWorksheet(classWithMastery.name.replace(/[*?:\\/\[\]]/g, "_"));

    worksheet.addRow([
      "Name",
      "Overall",
      ,
      "Limits and continuity",
      ,
      "Differentiation: definition and basic derivative rules",
      ,
      "Differentiation: composite, implicit, and inverse functions",
      ,
      "Contextual applications of differentiation",
      ,
      "Applying derivatives to analyze functions",
      ,
      "Integration and accumulation of change",
      ,
      "Differential equations",
      ,
      "Applications of integration",
    ]);

    worksheet.getRow(1).height = 35.05;

    const arial = {
      name: "Arial",
      size: 10,
      family: 2,
      charset: 1,
    };

    for (let i = 2; i <= 19; i++) {
      const col = worksheet.getColumn(i);
      if (i % 2 === 0) {
        col.numFmt = "0.0%";
      }

      col.font = { size: 11, name: "Calibri", family: 1, charset: 1 };
      col.width = 11.6;
    }

    worksheet.getColumn(1).width = 26.24;
    worksheet.getColumn(1).font = { ...arial };

    worksheet.getRow(1).numFmt = "General";
    worksheet.getRow(1).alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
      shrinkToFit: true,
      indent: 0,
      textRotation: 0,
    };
    worksheet.getRow(1).font = { ...arial, bold: true };

    worksheet.mergeCells("B1:C1");
    worksheet.mergeCells("D1:E1");
    worksheet.mergeCells("F1:G1");
    worksheet.mergeCells("H1:I1");
    worksheet.mergeCells("J1:K1");
    worksheet.mergeCells("L1:M1");
    worksheet.mergeCells("N1:O1");
    worksheet.mergeCells("P1:Q1");
    worksheet.mergeCells("R1:S1");

    worksheet.addConditionalFormatting({
      ref: "B1:B1048576 D1:D1048576 F1:F1048576 H1:H1048576 J1:J1048576 L1:L1048576 N1:N1048576 P1:P1048576 R1:R1048576 B1:S1",
      rules: [
        {
          type: "colorScale",
          priority: 2,
          cfvo: [
            { type: "num", value: 0 },
            { type: "num", value: 1 },
          ],
          color: [{ argb: "FFFFFFFF" }, { argb: "FF8577E2" }],
        },
      ],
    });

    let rowNumber = 2;
    for (const mastery of classWithMastery.masteries) {
      const totalPoints = mastery.units.reduce((a, { pointsAvailable: v }) => a + v, 0);
      const rowValues = [
        mastery.name,
        { formula: `C${rowNumber}/${totalPoints}` },
        {
          formula: `SUM(E${rowNumber},G${rowNumber},I${rowNumber},K${rowNumber},M${rowNumber},O${rowNumber},Q${rowNumber},S${rowNumber})`,
        },
      ] as any[];

      let columnNumber = 4;
      for (const topic of topicOrder) {
        const unit = mastery.units.find((t) => t.id === topic);

        if (unit) {
          rowValues.push(
            { formula: `${toLetter(columnNumber + 1)}${rowNumber}/${unit.pointsAvailable}` },
            unit.pointsEarned,
          );
        } else {
          rowValues.push("", "");
        }

        columnNumber += 2;
      }

      worksheet.addRow(rowValues);

      rowNumber++;
    }
  }

  const outFile = path.join(outDir, format(new Date(), "yyyyMMdd_HHmmss") + ".xlsx");

  await workbook.xlsx.writeFile(outFile);

  await page.evaluate((filename) => {
    window.setOverlay(`OK: Saved to ${filename}`);
  }, outFile);

  await sleep(2500);

  let countdown = 5;
  while (countdown >= 0) {
    await page.evaluate((countdown) => {
      window.setOverlay(`Closing in ${countdown} second(s)...`);
    }, countdown);

    await sleep(1000);
    countdown--;
  }

  await browser.close();
  process.exit(0);
}

main();
