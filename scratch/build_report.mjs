import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { Workbook, SpreadsheetFile, FileBlob } from '@oai/artifact-tool';

const root = process.cwd();
const data = JSON.parse(await fs.readFile(path.join(root, 'scratch/contributions.json'), 'utf8'));
const audit = JSON.parse(await fs.readFile(path.join(root, 'scratch/independent_audit/current_audit.json'), 'utf8'));
if (process.argv.includes('--add-overall-rank') || process.argv.includes('--apply-overall-tiebreak')) {
  const fourthRun = process.argv.includes('--apply-overall-tiebreak');
  const instructionHash = createHash('sha256').update(await fs.readFile(path.join(root, 'INSTRUCTIONS.md'))).digest('hex');
  const expectedInstructionHash = process.argv.find(argument => argument.startsWith('--instruction-sha256='))?.split('=')[1];
  if (!expectedInstructionHash || instructionHash !== expectedInstructionHash.toLowerCase()) throw new Error('Instructions changed since review.');
  const outputDir = path.join(root, 'outputs');
  const inventory = (await fs.readdir(outputDir, { recursive: true })).filter(name => name.toLowerCase().endsWith('.xlsx') && !path.basename(name).startsWith('~$')).sort();
  const filenames = ['January_Summary.xlsx', 'February_Summary.xlsx', 'March_Summary.xlsx', 'April_Summary.xlsx', 'Final_Summary.xlsx'];
  const monthlyHashes = {};
  let finalBook;
  let memberChecks = 0;
  for (const filename of filenames) {
    const bytes = await fs.readFile(path.join(outputDir, filename));
    if (filename !== 'Final_Summary.xlsx') monthlyHashes[filename] = createHash('sha256').update(bytes).digest('hex');
    const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path.join(outputDir, filename)));
    const summary = wb.worksheets.getItem('Summary');
    const isFinal = filename === 'Final_Summary.xlsx';
    const name = filename.replace('_Summary.xlsx', '');
    const rows = summary.getRange(isFinal ? 'A7:H25' : 'A7:F25').values;
    const fields = isFinal ? [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 6]] : [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 5]];
    const actualNames = rows.map(row => `${row[0]} ${row[1]}`);
    if (new Set(actualNames).size !== 19 || data.roster.some(member => !actualNames.includes(member))) throw new Error(`${filename}: source roster mismatch`);
    for (const row of rows) {
      const member = `${row[0]} ${row[1]}`;
      const expected = audit.members[member][isFinal ? 'final' : `${name} 2025`];
      for (const [field, column] of fields) {
        if (row[column] !== expected[field]) throw new Error(`${filename}/${member}/${field}: source mismatch`);
        memberChecks++;
      }
      if (isFinal && (row[5] !== expected.tithe_rank_combined || row[7] !== expected.offering_rank_combined)) throw new Error(`${member}: consistency rank mismatch`);
    }
    const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 20 }, maxChars: 2000 });
    if (!errors.ndjson.includes('matched 0 entries')) throw new Error(`${filename}: ${errors.ndjson}`);
    if (isFinal) finalBook = wb;
  }
  const summary = finalBook.worksheets.getItem('Summary');
  const beforeRows = summary.getRange(fourthRun ? 'A7:J25' : 'A7:H25').values;
  const beforeFormulas = summary.getRange(fourthRun ? 'A7:I25' : 'A7:H25').formulas;
  if (!fourthRun) {
  // Extend the existing last column's formatting without changing its content.
  for (const column of ['I', 'J']) {
    summary.getRange(`${column}1:${column}39`).format.font = { name: 'Arial', size: 11, color: '#172B45' };
    summary.getRange(`${column}6:${column}26`).copyFrom(summary.getRange('H6:H26'), 'all');
    summary.getRange(`${column}2`).copyFrom(summary.getRange('H2'), 'all');
  }
  summary.getRange('I:I').format.columnWidth = 14;
  summary.getRange('J:J').format.columnWidth = 17;
  summary.getRange('I6:J6').values = [['Average', 'Overall Rank']];
  summary.getRange('I6:J6').format = { fill: '#253C59', font: { name: 'Arial', size: 11, bold: true, color: '#FFFFFF' }, wrapText: true,
    horizontalAlignment: 'center', verticalAlignment: 'center', borders: { insideVertical: { style: 'thin', color: '#FFFFFF' } } };
  summary.getRange('I6').format.borders = { left: { style: 'thin', color: '#FFFFFF' }, right: { style: 'thin', color: '#FFFFFF' } };
  summary.getRange('I7:J25').format.horizontalAlignment = 'right';
  summary.getRange('I7:J25').format.verticalAlignment = 'center';
  for (let row = 8; row <= 24; row += 2) summary.getRange(`I${row}:J${row}`).format.fill = '#F6F8FB';
  summary.getRange('I26:J26').format.fill = '#EEF3F8';
  summary.getRange('I26:J26').format.borders = { top: { style: 'thin', color: '#B7C4D3' } };
  summary.getRange('I2:J2').format.borders = { bottom: { style: 'thin', color: '#B7C4D3' } };
  summary.getRange('I7:J25').formulas = beforeRows.map((_, index) => {
    const row = index + 7;
    return [`=AVERAGE(F${row},H${row})`, `=1+COUNTIF($I$7:$I$25,"<"&I${row})`];
  });
  summary.getRange('I7:I25').setNumberFormat('0.0');
  summary.getRange('J7:J25').setNumberFormat('0');
  summary.getRange('I26:J26').clear({ applyTo: 'contents' });
  summary.getRange('A32').copyFrom(summary.getRange('A31'), 'all');
  summary.getRange('A32').values = [['Average is the mean of the two consistency ranks. Lower averages rank higher; equal averages share an overall rank.']];
  summary.getRange('A32').format.font = { name: 'Arial', size: 10, italic: true, color: '#57697D' };
  } else {
    if (JSON.stringify(summary.getRange('I6:J6').values[0]) !== JSON.stringify(['Average', 'Overall Rank'])) throw new Error('Expected Average and Overall Rank columns are missing.');
    summary.getRange('J7:J25').formulas = beforeRows.map((_, index) => {
      const row = index + 7;
      return [`=1+COUNTIF($I$7:$I$25,"<"&I${row})+SUMPRODUCT(($I$7:$I$25=I${row})*(($C$7:$C$25+$D$7:$D$25)>(C${row}+D${row})))`];
    });
    summary.getRange('A32').values = [['Overall rank uses the lower average first, then the higher combined tithe and offering total for equal averages.']];
  }
  const rows = summary.getRange('A7:J25').values;
  const expectedAverages = beforeRows.map(row => (row[5] + row[7]) / 2);
  const combinedTotals = beforeRows.map(row => row[2] + row[3]);
  const results = rows.map((row, index) => {
    const expectedRank = 1 + expectedAverages.filter((value, other) => value < expectedAverages[index] || (fourthRun && value === expectedAverages[index] && combinedTotals[other] > combinedTotals[index])).length;
    if (row[8] !== expectedAverages[index] || row[9] !== expectedRank) throw new Error(`Average/overall rank mismatch for ${row[0]} ${row[1]}`);
    if (JSON.stringify(row.slice(0, fourthRun ? 9 : 8)) !== JSON.stringify(beforeRows[index].slice(0, fourthRun ? 9 : 8))) throw new Error('Existing member content changed.');
    return { member: `${row[0]} ${row[1]}`, titheRank: row[5], offeringRank: row[7], average: row[8], combinedTotal: combinedTotals[index], overallRank: row[9], ...(fourthRun ? { beforeOverallRank: beforeRows[index][9] } : {}) };
  });
  if (JSON.stringify(summary.getRange(fourthRun ? 'A7:I25' : 'A7:H25').formulas) !== JSON.stringify(beforeFormulas)) throw new Error('Existing formulas changed.');
  if (fourthRun && results.map(row => row.overallRank).sort((a, b) => a - b).some((rank, index) => rank !== index + 1)) throw new Error('Overall ranks are not unique 1–19; investigate an unresolved tie.');
  const totals = summary.getRange('C26:H26').values[0];
  if (totals[0] !== 33970 || totals[1] !== 6495 || totals[2] !== 62 || totals[4] !== 261) throw new Error('Final totals mismatch.');
  const errors = await finalBook.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 20 }, maxChars: 2000 });
  if (!errors.ndjson.includes('matched 0 entries')) throw new Error(errors.ndjson);
  const inspected = await finalBook.inspect({ kind: 'table', range: 'Summary!F6:J25', include: 'values,formulas', tableMaxRows: 20, tableMaxCols: 5, maxChars: 6500 });
  await fs.writeFile(path.join(root, fourthRun ? 'scratch/fourth_run_inspection.ndjson' : 'scratch/overall_rank_inspection.ndjson'), inspected.ndjson);
  const preview = await finalBook.render({ sheetName: 'Summary', range: 'A1:J39', scale: 1.2, format: 'png' });
  await fs.writeFile(path.join(root, fourthRun ? 'scratch/previews/Fourth_Run_Final.png' : 'scratch/previews/Overall_Rank_Final.png'), new Uint8Array(await preview.arrayBuffer()));
  for (const source of audit.files) {
    if (createHash('sha256').update(await fs.readFile(path.join(root, source.file))).digest('hex') !== source.sha256) throw new Error(`Original source changed: ${source.file}`);
  }
  if (createHash('sha256').update(await fs.readFile(path.join(root, 'INSTRUCTIONS.md'))).digest('hex') !== instructionHash) throw new Error('Instructions changed during validation.');
  if (!process.argv.includes('--preview-only')) {
    const output = await SpreadsheetFile.exportXlsx(finalBook);
    await output.save(path.join(outputDir, 'Final_Summary.xlsx'));
  }
  for (const [filename, expectedHash] of Object.entries(monthlyHashes)) {
    if (createHash('sha256').update(await fs.readFile(path.join(outputDir, filename))).digest('hex') !== expectedHash) throw new Error(`Monthly workbook changed: ${filename}`);
  }
  const afterInventory = (await fs.readdir(outputDir, { recursive: true })).filter(name => name.toLowerCase().endsWith('.xlsx') && !path.basename(name).startsWith('~$')).sort();
  if (JSON.stringify(inventory) !== JSON.stringify(afterInventory)) throw new Error('Spreadsheet inventory changed.');
  await fs.writeFile(path.join(root, fourthRun ? 'scratch/fourth_run_checks.json' : 'scratch/overall_rank_checks.json'), JSON.stringify({ saved: !process.argv.includes('--preview-only'), instructionHash, memberChecks, results, formulaErrors: 0, unchangedSpreadsheetInventory: true, monthlyWorkbooksByteIdentical: true, originalSourcesUnchanged: true, ...(fourthRun ? { overallRankRule: 'average ascending, combined tithe plus offering descending', uniqueOverallRanks: true } : {}) }, null, 2));
  console.log(`${process.argv.includes('--preview-only') ? 'Prepared preview of' : 'Updated existing'} Final_Summary.xlsx: ${fourthRun ? 'combined-payment tie-breaker in Overall Rank' : 'Average and Overall Rank'} for all 19 members; ${memberChecks} source value checks passed; monthly files and originals unchanged.`);
  process.exit(0);
}
if (process.argv.includes('--update-existing-combined')) {
  const outputDir = path.join(root, 'outputs');
  const logFile = path.join(root, 'logs/WORK_LOG.md');
  const instructionHash = createHash('sha256').update(await fs.readFile(path.join(root, 'INSTRUCTIONS.md'))).digest('hex');
  const expectedInstructionHash = process.argv.find(argument => argument.startsWith('--instruction-sha256='))?.split('=')[1];
  if (expectedInstructionHash && instructionHash !== expectedInstructionHash.toLowerCase()) throw new Error('Instructions changed since the last review. Reread the file before updating.');
  const filenames = ['January_Summary.xlsx', 'February_Summary.xlsx', 'March_Summary.xlsx', 'April_Summary.xlsx', 'Final_Summary.xlsx'];
  const inventory = (await fs.readdir(outputDir, { recursive: true })).filter(name => name.toLowerCase().endsWith('.xlsx') && !path.basename(name).startsWith('~$')).sort();
  const books = [];
  const numericRankChanges = [];
  let checkCount = 0;
  function match(actual, expected, label) {
    if (typeof actual !== 'number' || Math.abs(actual - expected) > 0.000001) throw new Error(`${label}: ${actual} != ${expected}`);
    checkCount++;
  }
  for (const filename of filenames) {
    const filePath = path.join(outputDir, filename);
    if (!(await fs.stat(filePath)).isFile()) throw new Error(`Existing output not found: ${filename}`);
    const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
    const summary = wb.worksheets.getItem('Summary');
    const isFinal = filename === 'Final_Summary.xlsx';
    const name = filename.replace('_Summary.xlsx', '');
    const rows = summary.getRange(isFinal ? 'A7:H25' : 'A7:F25').values;
    const fields = isFinal ? [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 6]] : [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 5]];
    const actualNames = rows.map(row => `${row[0]} ${row[1]}`);
    if (new Set(actualNames).size !== 19 || data.roster.some(member => !actualNames.includes(member))) throw new Error(`${filename}: roster differs from source`);
    for (const row of rows) {
      const member = `${row[0]} ${row[1]}`;
      const expected = audit.members[member][isFinal ? 'final' : `${name} 2025`];
      for (const [field, col] of fields) match(row[col], expected[field], `${filename}/${member}/${field}`);
    }
    if (isFinal) {
      // With the category amount already equal, comparing combined amounts
      // is equivalent to comparing the other category's amount.
      summary.getRange('F7:F25').formulas = rows.map((_, index) => {
        const row = index + 7;
        return [`=1+COUNTIF($E$7:$E$25,">"&E${row})+COUNTIFS($E$7:$E$25,E${row},$C$7:$C$25,">"&C${row})+COUNTIFS($E$7:$E$25,E${row},$C$7:$C$25,C${row},$D$7:$D$25,">"&D${row})`];
      });
      summary.getRange('H7:H25').formulas = rows.map((_, index) => {
        const row = index + 7;
        return [`=1+COUNTIF($G$7:$G$25,">"&G${row})+COUNTIFS($G$7:$G$25,G${row},$D$7:$D$25,">"&D${row})+COUNTIFS($G$7:$G$25,G${row},$D$7:$D$25,D${row},$C$7:$C$25,">"&C${row})`];
      });
      summary.getRange('A29').values = [['Remaining ties use combined tithe plus offering total. Higher combined amounts rank higher.']];
      const afterRows = summary.getRange('A7:H25').values;
      for (const [index, row] of afterRows.entries()) {
        const member = `${row[0]} ${row[1]}`;
        match(row[5], audit.members[member].final.tithe_rank_combined, `${member}/tithe rank`);
        match(row[7], audit.members[member].final.offering_rank_combined, `${member}/offering rank`);
        for (const [column, category] of [[5, 'tithe'], [7, 'offering']]) {
          if (row[column] !== rows[index][column]) numericRankChanges.push({ member, category, before: rows[index][column], after: row[column], combined: row[2] + row[3] });
        }
      }
      for (const column of [5, 7]) {
        const ranks = afterRows.map(row => row[column]).sort((a, b) => a - b);
        ranks.forEach((value, index) => match(value, index + 1, 'Unique ranks 1–19'));
      }
      const totals = summary.getRange('C26:H26').values[0];
      [[0, 33970], [1, 6495], [2, 62], [4, 261]].forEach(([column, value]) => match(totals[column], value, 'Final totals'));
      const monthlyTotals = summary.getRange('B39:E39').values[0];
      [33970, 6495, 62, 261].forEach((value, index) => match(monthlyTotals[index], value, 'Monthly-to-final reconciliation'));
      const inspected = await wb.inspect({ kind: 'table', range: 'Summary!A6:H26', include: 'values,formulas', tableMaxRows: 21, tableMaxCols: 8, maxChars: 7000 });
      await fs.writeFile(path.join(root, 'scratch/rerun_after_inspection.ndjson'), inspected.ndjson);
      const preview = await wb.render({ sheetName: 'Summary', range: 'A1:H39', scale: 1.2, format: 'png' });
      await fs.writeFile(path.join(root, 'scratch/previews/Rerun_After_Final.png'), new Uint8Array(await preview.arrayBuffer()));
    }
    const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 20 }, maxChars: 2000 });
    if (!errors.ndjson.includes('matched 0 entries')) throw new Error(`${filename}: ${errors.ndjson}`);
    books.push({ wb, filename, filePath });
  }
  // Check original source integrity before saving any output.
  for (const source of audit.files) {
    const bytes = await fs.readFile(path.join(root, source.file));
    if (createHash('sha256').update(bytes).digest('hex') !== source.sha256) throw new Error(`Source changed: ${source.file}`);
  }
  if (createHash('sha256').update(await fs.readFile(path.join(root, 'INSTRUCTIONS.md'))).digest('hex') !== instructionHash) throw new Error('Instructions changed while validating. Reread before saving.');
  for (const { wb, filename, filePath } of books) {
    if (process.argv.includes('--resume-final-only') && filename !== 'Final_Summary.xlsx') continue;
    const output = await SpreadsheetFile.exportXlsx(wb);
    await output.save(filePath);
    await fs.appendFile(logFile, `\n- Updated existing outputs/${filename} in place. No additional spreadsheet filename was created.\n`);
    console.log(`Updated existing ${filename}`);
  }
  const afterInventory = (await fs.readdir(outputDir, { recursive: true })).filter(name => name.toLowerCase().endsWith('.xlsx') && !path.basename(name).startsWith('~$')).sort();
  if (JSON.stringify(inventory) !== JSON.stringify(afterInventory)) throw new Error('Spreadsheet file inventory changed unexpectedly.');
  await fs.writeFile(path.join(root, 'scratch/rerun_checks.json'), JSON.stringify({ checks: checkCount, instructionHash, outputFiles: filenames, unchangedSpreadsheetInventory: true, allSourceFingerprintsUnchanged: true, numericRankChanges, rankRule: 'count descending, corresponding amount descending, combined tithe plus offering descending' }, null, 2));
  const rankChangeText = numericRankChanges.length ? numericRankChanges.map(change => `${change.member}: ${change.category} rank ${change.before} to ${change.after}, combined total ${change.combined}`).join('; ') : 'No numerical rank changes were required; the existing rankings already follow the saved three-stage rule.';
  await fs.appendFile(logFile, `\n- Checked ${checkCount} recalculated member values, ranking results, unique rank positions, and cumulative reconciliations. All five workbook formula-error scans found no errors. The spreadsheet inventory in outputs is unchanged. Original source fingerprints remain unchanged.\n- All 19 members have unique ranks 1–19 in both categories. No ties remain after applying combined totals.\n- Ranking results: ${rankChangeText}\n- Visual inspection and independent preservation checks are pending before completion.\n`);
  console.log(`Validated ${checkCount} recalculated values and rankings. No additional spreadsheet files created.`);
  process.exit(0);
}
if (process.argv.includes('--check-existing')) {
  const checks = [];
  const monthlyNames = ['January', 'February', 'March', 'April'];
  for (const name of [...monthlyNames, 'Final']) {
    const filename = `${name}_Summary.xlsx`;
    const filePath = path.join(root, 'outputs', filename);
    const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(filePath));
    const sheet = wb.worksheets.getItem('Summary');
    const values = sheet.getRange(name === 'Final' ? 'A7:H25' : 'A7:F25').values;
    const fields = name === 'Final' ? [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 6]] : [['tithe', 2], ['offering', 3], ['tithe_paid', 4], ['offering_paid', 5]];
    const actualNames = values.map(row => `${row[0]} ${row[1]}`);
    if (new Set(actualNames).size !== 19 || data.roster.some(member => !actualNames.includes(member))) throw new Error(`${filename}: roster differs from source`);
    for (const row of values) {
      const member = `${row[0]} ${row[1]}`;
      const expected = audit.members[member][name === 'Final' ? 'final' : `${name} 2025`];
      for (const [field, col] of fields) if (Math.abs(row[col] - expected[field]) > 0.000001) throw new Error(`${filename}/${member}/${field} mismatch`);
    }
    const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!', options: { useRegex: true, maxResults: 20 }, maxChars: 2000 });
    if (!errors.ndjson.includes('matched 0 entries')) throw new Error(`${filename}: ${errors.ndjson}`);
    const inspection = await wb.inspect({ kind: 'table', range: name === 'Final' ? 'Summary!A6:H10' : 'Summary!A6:F10', include: 'values,formulas', tableMaxRows: 5, tableMaxCols: 8, maxChars: 4000 });
    checks.push({ filename, matchesSource: true, formulaErrors: errors.ndjson, inspection: inspection.ndjson });
    if (name === 'Final') {
      const preview = await wb.render({ sheetName: 'Summary', range: 'A1:H39', scale: 1.2, format: 'png' });
      await fs.writeFile(path.join(root, 'scratch/previews/Rerun_Before_Final.png'), new Uint8Array(await preview.arrayBuffer()));
      await fs.writeFile(path.join(root, 'scratch/rerun_before_values.json'), JSON.stringify({ rows: values, formulas: sheet.getRange('A7:H25').formulas, note: sheet.getRange('A29').values }, null, 2));
    }
    console.log(`Rechecked existing ${filename}: all 19 members, amounts and counts match sources; no formula errors.`);
  }
  await fs.writeFile(path.join(root, 'scratch/rerun_checks.json'), JSON.stringify(checks, null, 2));
  process.exit(0);
}
const outputDir = path.join(root, 'outputs');
const previewDir = path.join(root, 'scratch/previews');
const logFile = path.join(root, 'logs/WORK_LOG.md');
const tieMode = process.argv.includes('--alphabetical-ties') ? 'alphabetical' : 'shared';
const months = ['January', 'February', 'March', 'April'];
const serial = iso => Date.parse(iso + 'T00:00:00Z') / 86400000 + 25569;
const num = '#,##0;(#,##0);0';
const colors = { ink: '#172B45', head: '#253C59', rule: '#B7C4D3', pale: '#EEF3F8', muted: '#57697D', source: '#17623E' };
const members = data.roster.map(member => {
  const parts = member.split(' ');
  if (parts.length !== 2) throw new Error(`Unexpected source name: ${member}`);
  return { member, first: parts[0], last: parts[1] };
});
if (members.length !== 19 || data.files.length !== 17 || data.records.length !== 323) throw new Error('Unexpected source coverage.');
await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(previewDir, { recursive: true });
let checks = 0;
const filesCreated = [];
const monthlyInputs = [];
const verification = [];
function equal(actual, expected, label) {
  if (typeof actual !== 'number' || Math.abs(actual - expected) > 0.000001) throw new Error(`${label}: ${actual} != ${expected}`);
  checks++;
}
function base(sheet, range, widths) {
  sheet.showGridLines = false;
  sheet.getRange(range).format.font = { name: 'Arial', size: 11, color: colors.ink };
  sheet.getRange(range).format.rowHeight = 23;
  sheet.getRange(range).format.verticalAlignment = 'center';
  widths.forEach((width, index) => sheet.getRange(`${String.fromCharCode(65 + index)}:${String.fromCharCode(65 + index)}`).format.columnWidth = width);
}
function top(sheet, title, context, last) {
  sheet.getRange('A1').values = [[title]];
  sheet.getRange('A1').format.font = { name: 'Arial', size: 18, bold: true, color: colors.ink };
  sheet.getRange(`A1:${last}1`).format.rowHeight = 32;
  sheet.getRange('A2').values = [[context]];
  sheet.getRange(`A2:${last}2`).format.borders = { bottom: { style: 'thin', color: colors.rule } };
}
function note(sheet, row, text) {
  sheet.getRange(`A${row}`).values = [[text]];
  sheet.getRange(`A${row}`).format.font = { name: 'Arial', size: 10, italic: true, color: colors.muted };
}
function header(sheet, range, height = 52) {
  sheet.getRange(range).format = { fill: colors.head, font: { name: 'Arial', size: 11, bold: true, color: '#FFFFFF' }, wrapText: true,
    horizontalAlignment: 'center', verticalAlignment: 'center', rowHeight: height,
    borders: { insideVertical: { style: 'thin', color: '#FFFFFF' } } };
}
function bands(sheet, first, last, col) {
  for (let row = first; row <= last; row++) if ((row - first) % 2) sheet.getRange(`A${row}:${col}${row}`).format.fill = '#F6F8FB';
}
function total(sheet, row, numericCols, last) {
  sheet.getRange(`A${row}`).values = [['Total']];
  for (const col of numericCols) sheet.getRange(`${col}${row}`).formulas = [[`=SUM(${col}7:${col}25)`]];
  sheet.getRange(`A${row}:${last}${row}`).format.fill = colors.pale;
  sheet.getRange(`A${row}:${last}${row}`).format.font.bold = true;
  sheet.getRange(`A${row}:${last}${row}`).format.borders = { top: { style: 'thin', color: colors.rule } };
  sheet.getRange(`C${row}:${last}${row}`).setNumberFormat(num);
}
async function validateAndExport(wb, filename, views) {
  const errors = await wb.inspect({ kind: 'match', searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
    options: { useRegex: true, maxResults: 30 }, maxChars: 2500 });
  if (!errors.ndjson.includes('matched 0 entries')) throw new Error(`${filename}: formula error scan: ${errors.ndjson}`);
  const inspected = await wb.inspect({ kind: 'table', range: 'Summary!A6:H10', include: 'values,formulas', tableMaxRows: 5, tableMaxCols: 8, maxChars: 4000 });
  verification.push({ file: filename, formulaErrors: errors.ndjson, inspected: inspected.ndjson });
  for (const [sheetName, range] of Object.entries(views)) {
    const blob = await wb.render({ sheetName, range, scale: 1.2, format: 'png' });
    await fs.writeFile(path.join(previewDir, `${filename.replace('.xlsx', '')}_${sheetName.replaceAll(' ', '_')}.png`), new Uint8Array(await blob.arrayBuffer()));
  }
  const output = await SpreadsheetFile.exportXlsx(wb);
  await output.save(path.join(outputDir, filename));
  filesCreated.push(filename);
  await fs.appendFile(logFile, `\n- Created outputs/${filename}. Reconciled all member amounts and payment counts against the fresh independent source audit. Formula error scan found no errors. Rendered both sheets for visual inspection.\n`);
  console.log(`Saved ${filename}`);
}

for (const month of months) {
  const monthNumber = months.indexOf(month) + 1;
  const monthKey = `2025-${String(monthNumber).padStart(2, '0')}-01`;
  const records = data.records.filter(record => record.month === monthKey);
  const sourceFiles = data.files.filter(file => file.month === monthKey);
  const wb = Workbook.create();
  const summary = wb.worksheets.add('Summary');
  const raw = wb.worksheets.add('Source records');
  const end = 4 + records.length;
  base(raw, `A1:I${end}`, [16, 18, 20, 15, 15, 15, 16, 36, 12]);
  top(raw, `${month} source records`, `Marvellous Church, ${month} 2025. Sheet1, rows 2–20 in each source file.`, 'I');
  raw.getRange('A4:I4').values = [['Service date', 'First Name', 'Last Name', 'Tithe', 'Offering', 'Tithe payment', 'Offering payment', 'Source file', 'Source row']];
  raw.getRange(`A5:I${end}`).values = records.map(record => {
    const [first, last] = record.member.split(' ');
    return [serial(record.date), first, last, record.tithe, record.offering, null, null, record.file, record.row];
  });
  raw.getRange(`F5:G${end}`).formulas = records.map((_, i) => {
    const row = i + 5;
    return [`=IF(AND(ISNUMBER(D${row}),D${row}>0),1,0)`, `=IF(AND(ISNUMBER(E${row}),E${row}>0),1,0)`];
  });
  header(raw, 'A4:I4');
  bands(raw, 5, end, 'I');
  raw.getRange(`A5:A${end}`).setNumberFormat('dd mmm yyyy');
  raw.getRange(`D5:G${end}`).setNumberFormat(num);
  raw.getRange(`D5:G${end}`).format.horizontalAlignment = 'right';
  raw.getRange(`I5:I${end}`).setNumberFormat('0');
  raw.getRange(`A5:E${end}`).format.font.color = colors.source;
  raw.getRange(`H5:I${end}`).format.font.color = colors.source;
  raw.tables.add(`A4:I${end}`, true, `${month}SourceRecords`).style = 'TableStyleMedium2';
  header(raw, 'A4:I4');
  raw.freezePanes.freezeRows(4);
  raw.freezePanes.freezeColumns(3);

  base(summary, 'A1:F33', [18, 22, 23, 23, 23, 23]);
  top(summary, `${month} contributions`, `Marvellous Church, ${month} 2025. ${sourceFiles.length} Sunday files and 19 members.`, 'F');
  note(summary, 4, 'Amounts use source units; currency is unspecified.');
  summary.getRange('A6:F6').values = [['First Name', 'Last Name', 'Total Tithe Paid', 'Total Offering Paid', 'Number of Tithe Payments', 'Number of Offering Payments']];
  summary.getRange('A7:B25').values = members.map(member => [member.first, member.last]);
  const rr = col => `'Source records'!$${col}$5:$${col}$${end}`;
  summary.getRange('C7:F25').formulas = members.map((_, i) => {
    const row = i + 7;
    return ['D', 'E', 'F', 'G'].map(col => `=SUMIFS(${rr(col)},${rr('B')},$A${row},${rr('C')},$B${row})`);
  });
  header(summary, 'A6:F6', 62);
  bands(summary, 7, 25, 'F');
  summary.getRange('C7:F26').setNumberFormat(num);
  summary.getRange('C7:F26').format.horizontalAlignment = 'right';
  total(summary, 26, ['C', 'D', 'E', 'F'], 'F');
  note(summary, 28, 'A positive amount counts once per member per Sunday, regardless of amount.');
  note(summary, 29, 'Blank means no payment recorded. Later-Sunday N/A tithes do not count as payments.');
  note(summary, 30, 'Tithe payments are scheduled on first Sundays. All 19 members are included.');
  note(summary, 31, 'Source records preserves original amounts, source filenames, and source row numbers.');
  const values = summary.getRange('C7:F25').values;
  for (let i = 0; i < members.length; i++) {
    const expected = audit.members[members[i].member][`${month} 2025`];
    ['tithe', 'offering', 'tithe_paid', 'offering_paid'].forEach((key, j) => equal(values[i][j], expected[key], `${month}/${members[i].member}/${key}`));
    monthlyInputs.push([serial(monthKey), members[i].first, members[i].last, ...values[i], `${month}_Summary.xlsx`]);
  }
  const monthlyTotal = summary.getRange('C26:F26').values[0];
  ['tithe', 'offering', 'tithe_paid', 'offering_paid'].forEach((key, i) => equal(monthlyTotal[i], audit.monthly[`${month} 2025`][key], `${month}/total/${key}`));
  for (const file of sourceFiles) {
    const fileRecords = records.filter(record => record.file === file.file);
    equal(fileRecords.reduce((sum, record) => sum + (typeof record.tithe === 'number' ? record.tithe : 0), 0), file.tithe, `${file.file}/tithe`);
    equal(fileRecords.reduce((sum, record) => sum + (typeof record.offering === 'number' ? record.offering : 0), 0), file.offering, `${file.file}/offering`);
  }
  await validateAndExport(wb, `${month}_Summary.xlsx`, { Summary: 'A1:F31', 'Source records': 'A1:I15' });
}

const wb = Workbook.create();
const summary = wb.worksheets.add('Summary');
const inputs = wb.worksheets.add('Monthly totals');
const inputEnd = 4 + monthlyInputs.length;
base(inputs, `A1:H${inputEnd}`, [15, 18, 21, 21, 23, 20, 21, 30]);
top(inputs, 'Monthly member totals', 'Marvellous Church, January–April 2025. Values from the four monthly summaries.', 'H');
inputs.getRange('A4:H4').values = [['Month', 'First Name', 'Last Name', 'Total Tithe Paid', 'Total Offering Paid', 'Tithe Payments', 'Offering Payments', 'Source summary']];
inputs.getRange(`A5:H${inputEnd}`).values = monthlyInputs;
header(inputs, 'A4:H4', 58);
bands(inputs, 5, inputEnd, 'H');
inputs.getRange(`A5:A${inputEnd}`).setNumberFormat('mmm yyyy');
inputs.getRange(`D5:G${inputEnd}`).setNumberFormat(num);
inputs.getRange(`D5:G${inputEnd}`).format.horizontalAlignment = 'right';
inputs.getRange(`A5:H${inputEnd}`).format.font.color = colors.source;
inputs.tables.add(`A4:H${inputEnd}`, true, 'MonthlyMemberTotals').style = 'TableStyleMedium2';
header(inputs, 'A4:H4', 58);
inputs.freezePanes.freezeRows(4);
inputs.freezePanes.freezeColumns(3);

base(summary, 'A1:H43', [18, 22, 26, 27, 22, 20, 22, 21]);
top(summary, 'Cumulative contributions', 'Marvellous Church, January–April 2025. 17 Sunday files and 19 members.', 'H');
note(summary, 4, 'Amounts use source units; currency is unspecified.');
summary.getRange('A6:H6').values = [['First Name', 'Last Name', 'Total Tithe Paid Across the Review Period', 'Total Offering Paid Across the Review Period', 'Number of Tithe Payments', 'Tithe Consistency Rank', 'Number of Offering Payments', 'Offering Consistency Rank']];
summary.getRange('A7:B25').values = members.map(member => [member.first, member.last]);
const ir = col => `'Monthly totals'!$${col}$5:$${col}$${inputEnd}`;
for (let i = 0; i < members.length; i++) {
  const row = i + 7;
  const sum = col => `SUMIFS(${ir(col)},${ir('B')},$A${row},${ir('C')},$B${row})`;
  const fallback = tieMode === 'alphabetical' ? `+COUNTIFS($E$7:$E$25,E${row},$C$7:$C$25,C${row},$A$7:$A$25,"<"&A${row})+COUNTIFS($E$7:$E$25,E${row},$C$7:$C$25,C${row},$A$7:$A$25,A${row},$B$7:$B$25,"<"&B${row})` : '';
  const offeringFallback = fallback.replaceAll('$E$', '$G$').replaceAll('$C$', '$D$').replaceAll(`E${row}`, `G${row}`).replaceAll(`C${row}`, `D${row}`);
  summary.getRange(`C${row}:H${row}`).formulas = [[
    `=${sum('D')}`, `=${sum('E')}`, `=${sum('F')}`,
    `=1+COUNTIF($E$7:$E$25,">"&E${row})+COUNTIFS($E$7:$E$25,E${row},$C$7:$C$25,">"&C${row})${fallback}`,
    `=${sum('G')}`,
    `=1+COUNTIF($G$7:$G$25,">"&G${row})+COUNTIFS($G$7:$G$25,G${row},$D$7:$D$25,">"&D${row})${offeringFallback}`
  ]];
}
header(summary, 'A6:H6', 78);
bands(summary, 7, 25, 'H');
summary.getRange('C7:H26').setNumberFormat(num);
summary.getRange('C7:H26').format.horizontalAlignment = 'right';
total(summary, 26, ['C', 'D', 'E', 'G'], 'H');
note(summary, 28, 'Each ranking uses payment count first, then the corresponding total amount for tied counts.');
note(summary, 29, tieMode === 'shared' ? 'Members tied on both count and amount share a rank; the following rank skips the tied positions.' : 'Exact ties on both count and amount use alphabetical name order to assign unique ranks.');
note(summary, 30, 'A positive amount counts once per Sunday. Blank and N/A entries do not count as payments.');
note(summary, 31, 'Monthly totals contains the values used from each monthly summary, without external links.');
summary.getRange('A34:E34').values = [['Month', 'Tithe total', 'Offering total', 'Tithe payments', 'Offering payments']];
summary.getRange('A35:A38').values = months.map((_, i) => [serial(`2025-${String(i + 1).padStart(2, '0')}-01`)]);
summary.getRange('B35:E38').formulas = months.map((_, i) => ['D', 'E', 'F', 'G'].map(col => `=SUMIF(${ir('A')},A${i + 35},${ir(col)})`));
header(summary, 'A34:E34', 48);
bands(summary, 35, 38, 'E');
summary.getRange('A35:A38').setNumberFormat('mmm yyyy');
summary.getRange('B35:E39').setNumberFormat(num);
summary.getRange('B35:E39').format.horizontalAlignment = 'right';
summary.getRange('A39').values = [['Period total']];
summary.getRange('B39:E39').formulas = [['=SUM(B35:B38)', '=SUM(C35:C38)', '=SUM(D35:D38)', '=SUM(E35:E38)']];
summary.getRange('A39:E39').format.fill = colors.pale;
summary.getRange('A39:E39').format.font.bold = true;
summary.getRange('A39:E39').format.borders = { top: { style: 'thin', color: colors.rule } };

const finalValues = summary.getRange('C7:H25').values;
for (let i = 0; i < members.length; i++) {
  const expected = audit.members[members[i].member].final;
  const fields = ['tithe', 'offering', 'tithe_paid', tieMode === 'shared' ? 'tithe_rank_shared' : 'tithe_rank', 'offering_paid', tieMode === 'shared' ? 'offering_rank_shared' : 'offering_rank'];
  fields.forEach((field, j) => equal(finalValues[i][j], expected[field], `${members[i].member}/final/${field}`));
  const monthlyRows = monthlyInputs.filter(row => row[1] === members[i].first && row[2] === members[i].last);
  for (const [summaryCol, inputCol] of [[0, 3], [1, 4], [2, 5], [4, 6]]) equal(finalValues[i][summaryCol], monthlyRows.reduce((sum, row) => sum + row[inputCol], 0), `${members[i].member}/monthly-to-final`);
}
const finalTotals = summary.getRange('C26:H26').values[0];
equal(finalTotals[0], 33970, 'Final tithe total');
equal(finalTotals[1], 6495, 'Final offering total');
equal(finalTotals[2], 62, 'Final tithe count');
equal(finalTotals[4], 261, 'Final offering count');
const periodTotals = summary.getRange('B39:E39').values[0];
[33970, 6495, 62, 261].forEach((value, i) => equal(periodTotals[i], value, 'Monthly-to-final reconciliation'));
await validateAndExport(wb, 'Final_Summary.xlsx', { Summary: 'A1:H39', 'Monthly totals': 'A1:H16' });

for (const source of audit.files) {
  const bytes = await fs.readFile(path.join(root, source.file));
  const fingerprint = createHash('sha256').update(bytes).digest('hex');
  if (fingerprint !== source.sha256) throw new Error(`Original source changed: ${source.file}`);
}
await fs.writeFile(path.join(root, 'scratch/verification.json'), JSON.stringify({ checks, tieMode, filesCreated, sourceFingerprintsUnchanged: true, verification }, null, 2));
await fs.appendFile(logFile, `\n## Completed calculation checks\n\n- Reviewed all 17 original Sunday files, covering every Sunday in January–April 2025. The same 19 members occur in each file.\n- Checked ${checks} member amounts, payment counts, ranks, file totals, and monthly-to-final reconciliations against independent calculations.\n- Final tithe: 33,970. Final offering: 6,495. Combined: 40,465 source units. Tithe payment occurrences: 62. Offering payment occurrences: 261.\n- All five formula-error scans found no spreadsheet errors. Every summary has 19 member rows with separate first and last names.\n- Verified all 17 source SHA-256 fingerprints against the independent audit after export; the source files remain unchanged.\n- Exact ties: Lawal Lawal and Obafemi Dolapo each have 4 tithe payments / 2,000; Alaye Zainab and Solomon Joseph each have 3 / 600. Ayomide Adeola and Hannah Ladele each have 13 offering payments / 235.\n- Exact-tie handling: ${tieMode === 'shared' ? 'shared competition ranks, with the next rank skipping tied positions. This assumption applies only when both specified ranking criteria are identical.' : 'alphabetical order after equal counts and equal total amounts to produce unique ranks 1–19.'}\n- The context uses the word quarter for January–April. The outputs use the explicitly requested four-month review period and include April.\n- Visual review of rendered sheets is the remaining final check.\n\n## Source integrity record\n\n| Source file | SHA-256 |\n|---|---|\n${audit.files.map(source => `| ${source.file.replaceAll('\\', '/')} | ${source.sha256} |`).join('\n')}\n`);
console.log(`Verified ${checks} calculations and ranks. All original source fingerprints unchanged. Tie mode: ${tieMode}.`);
