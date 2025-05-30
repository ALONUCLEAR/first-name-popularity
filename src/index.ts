import * as xlsx from 'xlsx';

export type XlsxReadOptions = {
    sheetsToInclude?: string[],
    startRow?: number;
    startColumn?: number;
    endColumn?: number;
};

export function xlsxToObject<T = string>(
    fileName: string,
    options: XlsxReadOptions = {}
): Record<string, T[]> {
    const workbook = xlsx.readFile(fileName);
    const worksheets: Record<string, T[]> = {};

    workbook.SheetNames.forEach(sheetname => {
        if ((options?.sheetsToInclude?.length ?? 0) > 0 && !options.sheetsToInclude!.includes(sheetname)) {
            return;
        }

        const sheet = workbook.Sheets[sheetname];

        if (!sheet) {
            return;
        }

        // Determine the range based on options
        const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1'); // Get the default range of the sheet

        const sRow = options.startRow ?? 0;
        const sCol = options.startColumn ?? 0;
        const eCol = options.endColumn ?? range.e.c; // Use sheet's last column if endColumn is not specified

        // Adjust the range for sheet_to_json
        const newRange = {
            s: { r: sRow, c: sCol },
            e: { r: range.e.r, c: eCol }
        };

        worksheets[sheetname] = xlsx.utils.sheet_to_json<T>(sheet, { range: newRange });
    });

    return worksheets;
}

const dataPath = 'src/assets/first-names.xlsx';

const sheetNamesToInclude = ["נשים יהודיות"];
const options: XlsxReadOptions = {
    sheetsToInclude: sheetNamesToInclude,
    startRow: 6,
    startColumn: 0,
    endColumn: 1
};

type RawData = { prati1: string,"1948-2023": number};
const rawData: Record<string, RawData[]> = xlsxToObject<RawData>(dataPath, options);

const populationPerName = rawData[sheetNamesToInclude[0]].map(({ prati1: name, "1948-2023": amount }) => ({ name, amount }));

const groupByFirstLetter: Record<string, number> = {};

for (const { name, amount } of populationPerName) {
    const firstLetter = name.charAt(0).toLocaleLowerCase();

    if(!groupByFirstLetter[firstLetter]) {
        groupByFirstLetter[firstLetter] = amount;
    } else {
        groupByFirstLetter[firstLetter] += amount;
    }
}

const letterPopularityRank = Object.entries(groupByFirstLetter)
  .sort((a, b) => b[1] - a[1])
  .map((data, index) => ({ rank: index + 1, name: data[0], amount: data[1] }));

console.dir(letterPopularityRank, { depth: Infinity });