const fs = require("node:fs");
const path = require("node:path");
const { randomUUID } = require("node:crypto");
const { pairRounds, matchPoints } = require("./compare");

function validDate(value) {
  return (
    typeof value === "string" &&
    /^\d{4}-\d{2}-\d{2}$/.test(value) &&
    Number.isFinite(Date.parse(value)) &&
    new Date(value).toISOString().startsWith(value)
  );
}
class Store {
  constructor(directory) {
    this.dir = directory;
    fs.mkdirSync(directory, { recursive: true });
  }
  filename(id) {
    if (!/^[a-f0-9-]{36}$/.test(id)) throw Error("รหัสงานไม่ถูกต้อง");
    return path.join(this.dir, id + ".json");
  }
  save(job) {
    const file = this.filename(job.id);
    const temp = file + "." + randomUUID() + ".tmp";
    try {
      fs.writeFileSync(temp, JSON.stringify(job), { flag: "wx", flush: true });
      fs.renameSync(temp, file);
    } finally {
      fs.rmSync(temp, { force: true });
    }
    return job;
  }
  load(id) {
    return JSON.parse(fs.readFileSync(this.filename(id), "utf8"));
  }
  list() {
    return fs
      .readdirSync(this.dir)
      .filter((n) => n.endsWith(".json"))
      .map((n) => {
        const { id, before, after, created } = this.load(n.slice(0, -5));
        return {
          id,
          plot: before.plot,
          from: before.year,
          to: after.year,
          created,
        };
      })
      .sort((a, b) => b.created.localeCompare(a.created));
  }
  create(before, after) {
    pairRounds(before, after);
    return this.save({
      id: randomUUID(),
      created: new Date().toISOString(),
      before,
      after,
      history: [],
    });
  }
  review(id, key, input) {
    const job = this.load(id);
    const point = matchPoints(job.before, job.after).find((p) => p.key === key);
    if (!point) throw Error("ไม่พบจุดปลูก");
    const { decision, note, checked, planted } = input;
    if (
      !["same", "replacement", "checked", "pending"].includes(decision) ||
      !validDate(checked) ||
      typeof note !== "string" ||
      (planted && !validDate(planted))
    )
      throw Error("กรุณาตรวจผลและวันที่อีกครั้ง");
    const paired =
      point.plot &&
      point.no &&
      point.before.length === 1 &&
      point.after.length === 1;
    if (!paired && ["same", "replacement"].includes(decision))
      throw Error("ต้องมีแถวคู่เดียวก่อนยืนยันตัวต้น");
    const event = {
      key,
      decision,
      note,
      checked,
      saved: new Date().toISOString(),
    };
    if (decision === "replacement") {
      event.planted = planted || null;
      event.previousLine = point.before[0].line;
      event.baselineLine = point.after[0].line;
      event.previousSheet = point.before[0].sourceSheet || job.before.sheet;
      event.baselineSheet = point.after[0].sourceSheet || job.after.sheet;
    }
    job.history.push(event);
    return this.save(job);
  }
}
module.exports = { Store };
