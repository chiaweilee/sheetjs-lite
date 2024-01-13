const fs = require("fs");
const path = require("path");
const full = require("../src/xlsx.full");

function getFixtures(ext) {
  const fixturesPath = path.join(__dirname, "./fixtures");
  const fixtures = [];
  fs.readdirSync(fixturesPath).forEach(function (filename) {
    if (!filename.endsWith(`.${ext}`)) return;
    const filePath = path.join(fixturesPath, filename);
    const filePathStat = fs.statSync(filePath);
    if (filePathStat.isFile()) {
      const content = fs.readFileSync(filePath);
      fixtures.push(content);
    }
  });
  return fixtures;
}

const xlsxFixtures = getFixtures("xlsx");
const xlsFixtures = getFixtures("xls");

describe("xlsx.full.js", () => {
  it(`parse XLSX (${xlsxFixtures.length} files)`, function () {
    const target = xlsxFixtures;
    expect.assertions(target.length * 5);

    target.forEach((data) => {
      const result = full.read(data);
      expect(typeof result).toBe("object");
      expect(typeof result.SheetNames).toBe("object");
      expect(typeof result.Sheets).toBe("object");
      expect(result.SheetNames.length).not.toBe(0);
      expect(result.SheetNames.length).toEqual(
        Object.keys(result.Sheets).length
      );
    });
  });

  it(`parse XLS (${xlsFixtures.length} files)`, function () {
    const target = xlsFixtures;
    expect.assertions(target.length * 5);
  
    target.forEach((data) => {
      const result = full.read(data);
      expect(typeof result).toBe("object");
      expect(typeof result.SheetNames).toBe("object");
      expect(typeof result.Sheets).toBe("object");
      expect(result.SheetNames.length).not.toBe(0);
      expect(result.SheetNames.length).toEqual(Object.keys(result.Sheets).length);
    });
  });
});
