import ExcelJS from "exceljs";

const COLUMNS_KEY = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
  "AA",
  "AB",
  "AC",
  "AD",
  "AE",
  "AF",
  "AG",
  "AH",
  "AI"
];

interface ledgerEntry {
  type: accountType;
  amount: number;
  formula: string;
  accountName?: string;
}

class Account {
  private _name: accountName;
  private _entries: ledgerEntry[] = [];
  private _type: accountType;
  private _debitColumn = "";
  private _creditColumn = "";

  constructor(name: accountName, type: accountType) {
    this._name = name;
    this._type = type;
  }

  get name() {
    return this._name;
  }

  get entries() {
    return this._entries;
  }

  get type() {
    return this._type;
  }

  addEntry(newEntry: ledgerEntry) {
    this._entries.push(newEntry);
  }

  setColumns(debitColumn: string, creditColumn: string) {
    this._debitColumn = debitColumn;
    this._creditColumn = creditColumn;
  }

  get debitColumn() {
    return this._debitColumn;
  }

  get creditColumn() {
    return this._creditColumn;
  }
}

interface transaction {
  account: string;
  amount: number;
  cell?: string;
}

interface entry {
  date?: object;
  debit?: transaction;
  credit?: transaction;
}

type accountType = "Debit" | "Credit";

type accountName =
  | "Cash harian"
  | "Cash"
  | "Stock"
  | "Omzet"
  | "Operational"
  | "Pengeluaran personal"
  | "Utang dagang"
  | "Salary"
  | "Piutang dagang"
  | "Utility"
  | "Pengeluaran rutin";

async function main() {
  const args = process.argv.slice(2);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("../sample2.xlsx");
  const journal = workbook.getWorksheet("Journal");

  const allEntries: entry[] = [];
  const totalEntryCount = journal.actualRowCount / 2;
  let rowIndex = 1;

  while (allEntries.length !== totalEntryCount) {
    const debitIndex = rowIndex;
    const debitRow = journal.getRow(debitIndex);
    const creditIndex = (rowIndex += 1);
    const creditRow = journal.getRow(creditIndex);

    const entry: entry = {
      date: Object(debitRow.getCell("A").value),
      debit: {
        account: String(debitRow.getCell("B").value),
        amount: Number(debitRow.getCell("D").toString()),
        cell: `D${debitIndex}`
      },
      credit: {
        account: String(creditRow.getCell("C").value),
        amount: Number(creditRow.getCell("E").toString()),
        cell: `E${creditIndex}`
      }
    };

    allEntries.push(entry);
    rowIndex += 2;
  }
  //console.log(allEntries);

  const allAccountNames: string[] = [];

  allEntries.forEach(entry => {
    if (entry.debit) allAccountNames.push(entry.debit.account);
    if (entry.credit) allAccountNames.push(entry.credit.account);
  });

  const distinctAccountNames = allAccountNames.filter(
    (n, i) => allAccountNames.indexOf(n) === i
  );
  console.log(distinctAccountNames);

  // initialize all accounts
  const AccountCashHarian = new Account("Cash harian", "Debit");
  const AccountCash = new Account("Cash", "Debit");
  const AccountStock = new Account("Stock", "Debit");
  const AccountOmzet = new Account("Omzet", "Credit");
  const AccountOperational = new Account("Operational", "Debit");
  const AccountPengeluaranPersonal = new Account(
    "Pengeluaran personal",
    "Debit"
  );
  const AccountUtangDagang = new Account("Utang dagang", "Credit");
  const AccountSalary = new Account("Salary", "Debit");
  const AccountPiutangDagang = new Account("Piutang dagang", "Credit");
  const AccountUtility = new Account("Utility", "Debit");
  const AccountPengeluaranRutin = new Account("Pengeluaran rutin", "Debit");

  const allAccounts: Account[] = [
    AccountCashHarian,
    AccountCash,
    AccountStock,
    AccountOmzet,
    AccountOperational,
    AccountPengeluaranPersonal,
    AccountUtangDagang,
    AccountSalary,
    AccountPiutangDagang,
    AccountUtility,
    AccountPengeluaranRutin
  ];

  // map thru all entries to respective account
  allEntries.forEach((entry, index) => {
    const ledgerEntries: ledgerEntry[] = [];
    if (entry.debit) {
      const ledgerDebitEntry: ledgerEntry = {
        type: "Debit",
        amount: entry.debit.amount,
        formula: `Journal!${entry.debit.cell}`,
        accountName: entry.debit.account
      };
      ledgerEntries.push(ledgerDebitEntry);
    }
    if (entry.credit) {
      const ledgerCreditEntry: ledgerEntry = {
        type: "Credit",
        amount: entry.credit.amount,
        formula: `Journal!${entry.credit.cell}`,
        accountName: entry.credit.account
      };
      ledgerEntries.push(ledgerCreditEntry);
    }
    ledgerEntries.forEach(ledgerEntry => {
      const targetAccount = allAccounts.find(
        acc => acc.name === ledgerEntry.accountName
      );

      if (targetAccount) targetAccount.addEntry(ledgerEntry);
      else throw Error("invalid account");
    });
  });

  // create new worksheet to write the ledger
  const ledger = workbook.addWorksheet("Ledger");
  //const ledger = workbook.getWorksheet("Ledger");

  // constructing account table header
  let a = 0;
  for (let x = 0; x < distinctAccountNames.length; x++) {
    const name = distinctAccountNames[x];
    const cell = ledger.getCell(`${COLUMNS_KEY[a]}1`);
    cell.value = name;
    cell.font = { bold: true };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    ledger.mergeCells(`${COLUMNS_KEY[a]}1:${COLUMNS_KEY[a + 1]}1`);

    const targetAccount = allAccounts.find(acc => acc.name === name);

    if (targetAccount)
      targetAccount.setColumns(COLUMNS_KEY[a], COLUMNS_KEY[a + 1]);
    else throw Error("account error >> cannot set columns");

    a += 3;
  }

  // writing the entries to respective account table

  allAccounts.forEach(account => {
    account.entries.forEach((entry, index) => {
      switch (entry.type as accountType) {
        case "Debit":
          const debitCell = ledger.getCell(
            `${account.debitColumn}${index + 2}`
          );
          debitCell.value = {
            ...debitCell,
            formula: entry.formula,
            result: entry.amount
          };
          break;

        case "Credit":
          const CreditCell = ledger.getCell(
            `${account.creditColumn}${index + 2}`
          );
          CreditCell.value = {
            ...CreditCell,
            formula: entry.formula,
            result: entry.amount
          };
          break;

        default:
          break;
      }

      // calculate total
      if (account.entries.length === index + 1) {
        const cellCode =
          account.type === "Debit"
            ? `${account.debitColumn}${index + 3}`
            : `${account.creditColumn}${index + 3}`;

        const formulaTotal =
          account.type === "Debit"
            ? `SUM((${account.debitColumn}2:${account.debitColumn}${index +
                2})-(${account.creditColumn}2:${account.creditColumn}${index +
                2}))`
            : `SUM((${account.creditColumn}2:${account.creditColumn}${index +
                2})-(${account.debitColumn}2:${account.debitColumn}${index +
                2}))`;

        const totalCell = ledger.getCell(cellCode);

        totalCell.font = { bold: true };
        totalCell.value = {
          ...totalCell,
          formula: formulaTotal,
          result: 3253000
        };
      }
    });
  });

  // save in the same file
  await workbook.xlsx.writeFile("../sample2.xlsx");
}

main();
