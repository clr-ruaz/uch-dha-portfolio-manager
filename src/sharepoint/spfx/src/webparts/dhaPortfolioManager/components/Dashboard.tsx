import * as React from "react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDashboardProps } from "./IDashboardProps";
import styles from "./Dashboard.module.scss";

interface Person {
  Id: number;
  Title: string;
  personName?: string;
  shortName?: string;
  propertyName?: string;
  unitName?: string;
  residencyStatus?: string;
  leaseStartDate?: string;
  leaseEndDate?: string;
  moveInDate?: string;
  moveOutDate?: string;
  prospectSourceName?: string;
  [key: string]: any;
}
interface Intake {
  Id: number;
  Title: string;
  TenantNameId?: number;
  TenantName?: Person;
  [key: string]: any;
}
interface LedgerEntry {
  Id: number;
  Title: string;
  TenantNameId?: number;
  date?: string;
  transactionType?: string;
  CategoryName?: string;
  Description?: string;
  Amount?: number;
  PaymentMethod?: string;
  DateReversed?: string;
}
interface LedgerPivotRow {
  key: string;
  tenant: string;
  property: string;
  unit: string;
  balance?: number;
  monthKey: string;
  month: string;
  charges: { [category: string]: number };
  payments: { [category: string]: number };
  credits: { [category: string]: number };
  transactions: LedgerEntry[];
}
interface LedgerPivot {
  chargeCategories: string[];
  paymentCategories: string[];
  creditCategories: string[];
  rows: LedgerPivotRow[];
}
type DataView = "resident" | "ledger";
type LedgerDetailSortKey =
  | "date"
  | "transactionType"
  | "CategoryName"
  | "Description"
  | "Amount"
  | "PaymentMethod"
  | "DateReversed";
type DataType = "text" | "money" | "date" | "link" | "action";
interface Column {
  key: string;
  label: string;
  type: DataType;
  value: (item: Intake) => any;
}
interface FormField {
  key: string;
  label: string;
  type: "text" | "money" | "date" | "boolean" | "status";
  readOnly?: boolean;
}
interface SnapshotMetric {
  label: string;
  value: number;
  color: string;
}
interface ResidentMetricGroup {
  label: string;
  segments: SnapshotMetric[];
}
interface PropertyMetric {
  name: string;
  residents: number;
  rent: number;
}
const money = new Intl.NumberFormat("en-US", {
  style: "currency",
  currency: "USD",
});
const pageSizeOptions = [10, 20, 50, 100];
const configurationListName = "DHA Configuration";
const refreshBalancesConfigurationTitle = "Refresh Balances Endpoint";
const llmRequestEndpointConfigurationTitle = "LLM Request Endpoint";
const summarizeTransactionsEndpointConfigurationTitle =
  "Summarize Transactions Endpoint";
const analyzeBalanceInstructionsConfigurationTitle =
  "Analyze Balance Instructions";
const configurationValueFieldKey = "Value";
const flowServiceResource = "https://service.flow.microsoft.com/";
const managePermissionsMask = 1 << 25;
const documentStatusOptions = [
  "Contract not uploaded",
  "Welcome/Invoice not uploaded",
  "Amendment not uploaded",
];
const documentStatusFields: { [status: string]: string[] } = {
  "Contract not uploaded": ["HAPContract"],
  "Welcome/Invoice not uploaded": [
    "WelcomeLetter",
    "WelcomeInvoice",
    "Welcome_x002f_Invoice",
    "Welcome_x0020_Invoice",
    "Invoice",
  ],
  "Amendment not uploaded": [
    "Amendment",
    "HAPAmendment",
    "HAP_x0020_Amendment",
  ],
};
const defaultResidency = [
  "Approved",
  "Approve",
  "Pending",
  "Current",
  "Notice to Vacate",
];
const defaultDha = ["Active", "Unknown"];
const fieldGroups: Array<{ title: string; fields: FormField[] }> = [
  {
    title: "Ledger Summary",
    fields: [
      { key: "Charges", label: "Charges", type: "money", readOnly: true },
      { key: "Payments", label: "Payment", type: "money", readOnly: true },
      { key: "Credits", label: "Credits", type: "money", readOnly: true },
      { key: "Balance", label: "Balance", type: "money", readOnly: true },
    ],
  },
  {
    title: "Lease Details",
    fields: [
      { key: "MonthlyRent", label: "Monthly Rent", type: "money" },
      { key: "HAPPortion", label: "HAP Portion", type: "money" },
      { key: "ApplicationFee", label: "Application Fee", type: "money" },
      { key: "AdministrativeFee", label: "Administrative Fee", type: "money" },
      { key: "RiskFee", label: "Risk Fee", type: "money" },
      { key: "SecurityDeposit", label: "Security Deposit", type: "money" },
      {
        key: "LandlordIncentiveFee",
        label: "Landlord Incentive Fee",
        type: "money",
      },
    ],
  },
  {
    title: "DHA Program",
    fields: [
      { key: "CaseworkerName", label: "Caseworker Name", type: "text" },
      { key: "CaseworkerContact", label: "Caseworker Contact", type: "text" },
      { key: "HAPContractStart", label: "HAP Start", type: "date" },
      { key: "HAPContractEnd", label: "HAP End", type: "date" },
      { key: "HAPAmendment1Date", label: "HAP Amend 1 Date", type: "date" },
      { key: "HAPAmendment1Portion", label: "HAP Portion 1", type: "money" },
      { key: "HAPAmendment2Date", label: "HAP Amend 2 Date", type: "date" },
      { key: "HAPAmendment2Portion", label: "HAP Portion 2", type: "money" },
      { key: "HAPAmendment3Date", label: "HAP Amend 3 Date", type: "date" },
      { key: "HAPAmendment3Portion", label: "HAP Portion 3", type: "money" },
      { key: "DHAStatus", label: "DHA Status", type: "status" },
    ],
  },
];
const fieldMap: { [key: string]: FormField } = {};
fieldGroups.forEach((group) =>
  group.fields.forEach((field) => {
    fieldMap[field.key] = field;
  })
);
const safe = (value: string): string => value.replace(/'/g, "''");
const name = (item: Intake): string =>
  (item.TenantName &&
    (item.TenantName.personName ||
      item.TenantName.shortName ||
      item.TenantName.Title)) ||
  item.Title ||
  "Unnamed resident";
const prop = (item: Intake): string =>
  (item.TenantName && item.TenantName.propertyName) ||
  item.Tenant_x0020_Name_x003a__x0020_p ||
  "";
const unit = (item: Intake): string =>
  (item.TenantName && item.TenantName.unitName) ||
  item.Tenant_x0020_Name_x003a__x0020_u ||
  "";
const residency = (item: Intake): string =>
  (item.TenantName && item.TenantName.residencyStatus) ||
  item.Tenant_x0020_Name_x003a__x0020_r ||
  "Unknown";
const dha = (item: Intake): string => item.DHAStatus || "Unknown";
const num = (value: any): number => Number(value || 0);
const date = (value: any): string =>
  value ? new Date(value).toLocaleDateString("en-US") : "—";
const dateInput = (value: any): string =>
  value ? String(value).slice(0, 10) : "";
const domainIdentifier = (value: any): string | undefined => {
  if (typeof value !== "string") return undefined;
  const identifier = value.trim();
  return identifier && !/^\d+$/.test(identifier) ? identifier : undefined;
};
const propertyGuid = (value: any): string | undefined => {
  if (typeof value !== "string") return undefined;
  const identifier = value.trim();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    identifier
  )
    ? identifier
    : undefined;
};
const residentPropertyId = (item: Intake): string | number | undefined => {
  const person = item.TenantName;
  const propertyId =
    item.propertyId ??
    item.PropertyID ??
    item.PropertyId ??
    person?.propertyId ??
    person?.PropertyID ??
    person?.PropertyId;
  return propertyId === undefined || propertyId === null || propertyId === ""
    ? undefined
    : propertyId;
};
const residentBillingAccountId = (item: Intake): string | number | undefined => {
  const person = item.TenantName;
  const billingAccountId =
    item.BillingAccountId ??
    item.billingAccountId ??
    item.BillingAccountID ??
    person?.BillingAccountId ??
    person?.billingAccountId ??
    person?.BillingAccountID;
  return billingAccountId === undefined ||
    billingAccountId === null ||
    billingAccountId === ""
    ? undefined
    : billingAccountId;
};
const normalizeMarkdown = (value: string): string => {
  let markdown = value.trim();
  try {
    const decoded = JSON.parse(markdown);
    if (typeof decoded === "string") markdown = decoded;
  } catch {
    // The endpoint normally returns Markdown directly.
  }
  return markdown.replace(/\\r\\n|\\n|\\r/g, "\n").replace(/\r\n?/g, "\n").trim();
};
const markdownInline = (value: string): React.ReactNode[] =>
  value.split(/(\*\*[^*]+\*\*)/g).map((part, index) =>
    /^\*\*[^*]+\*\*$/.test(part) ? (
      <strong key={index}>{part.slice(2, -2)}</strong>
    ) : (
      part
    )
  );
const isMarkdownTableSeparator = (line: string): boolean =>
  /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)+\|?\s*$/.test(line);
const markdownTableCells = (line: string): string[] => {
  const trimmed = line.trim().replace(/^\|/, "").replace(/\|$/, "");
  return trimmed.split("|").map((cell) => cell.trim());
};
const renderMarkdown = (value: string): React.ReactNode[] => {
  const lines = value.replace(/\r\n?/g, "\n").split("\n");
  const rendered: React.ReactNode[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    if (
      index + 1 < lines.length &&
      line.indexOf("|") >= 0 &&
      isMarkdownTableSeparator(lines[index + 1])
    ) {
      const headers = markdownTableCells(line);
      const rows: string[][] = [];
      index += 2;
      while (
        index < lines.length &&
        lines[index].indexOf("|") >= 0 &&
        lines[index].trim()
      ) {
        rows.push(markdownTableCells(lines[index]));
        index += 1;
      }
      rendered.push(
        <div key={`table-${index}`} className={styles.analysisTableWrap}>
          <table className={styles.analysisTable}>
            <thead>
              <tr>
                {headers.map((header, headerIndex) => {
                  const isAmountColumn =
                    header.replace(/\*/g, "").trim().toLowerCase() === "amount";
                  return (
                    <th
                      key={headerIndex}
                      className={isAmountColumn ? styles.amountColumn : undefined}
                    >
                      {markdownInline(header)}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {headers.map((header, cellIndex) => (
                    <td
                      key={cellIndex}
                      className={
                        header.replace(/\*/g, "").trim().toLowerCase() === "amount"
                          ? styles.amountColumn
                          : undefined
                      }
                    >
                      {markdownInline(row[cellIndex] || "")}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
      index -= 1;
      continue;
    }
    const heading = line.match(/^(#{1,3})\s+(.+)$/);
    const unorderedItem = line.match(/^[-*]\s+(.+)$/);
    const orderedItem = line.match(/^\d+\.\s+(.+)$/);
    if (heading) {
      const Heading =
        heading[1].length === 1
          ? "h3"
          : heading[1].length === 2
          ? "h4"
          : "h5";
      rendered.push(<Heading key={index}>{markdownInline(heading[2])}</Heading>);
      continue;
    }
    if (unorderedItem) {
      rendered.push(
        <p key={index} className={styles.analysisListItem}>
          • {markdownInline(unorderedItem[1])}
        </p>
      );
      continue;
    }
    if (orderedItem) {
      rendered.push(
        <p key={index} className={styles.analysisListItem}>
          {line.match(/^\d+\./)![0]} {markdownInline(orderedItem[1])}
        </p>
      );
      continue;
    }
    if (!line.trim()) {
      rendered.push(<div key={index} className={styles.analysisSpacer} />);
      continue;
    }
    rendered.push(<p key={index}>{markdownInline(line)}</p>);
  }

  return rendered;
};
const comparableFieldValue = (
  field: FormField,
  value: any
): string | boolean => {
  if (field.type === "boolean") return Boolean(value);
  if (field.type === "date") return dateInput(value);
  if (field.type === "money")
    return value === undefined || value === null || value === ""
      ? ""
      : String(Number(value));
  return value === undefined || value === null ? "" : String(value);
};
const docUrl = (base: string, item: Intake): string =>
  `${base}/Shared Documents/${encodeURIComponent(
    `${name(item)} (${item.Id})`
  )}/`;
const csvCell = (value: string): string => {
  const safeValue = /^[=+\-@]/.test(value) ? `'${value}` : value;
  return `"${safeValue.replace(/"/g, '""')}"`;
};
const hasDocumentValue = (value: any): boolean => {
  if (value === undefined || value === null || value === "") return false;
  if (typeof value === "string") return Boolean(value.trim());
  if (Array.isArray(value)) return value.some(hasDocumentValue);
  if (typeof value === "object")
    return Object.keys(value).some(
      (key) => key !== "__metadata" && hasDocumentValue(value[key])
    );
  return Boolean(value);
};
const matchesDocumentStatus = (item: Intake, status: string): boolean =>
  (documentStatusFields[status] || []).every(
    (field) => !hasDocumentValue(item[field])
  );
const col: Column[] = [
  { key: "property", label: "Property", type: "text", value: prop },
  { key: "unit", label: "Unit", type: "text", value: unit },
  { key: "charges", label: "Charges", type: "money", value: (r) => r.Charges },
  { key: "payment", label: "Payment", type: "money", value: (r) => r.Payments },
  { key: "credits", label: "Credits", type: "money", value: (r) => r.Credits },
  { key: "balance", label: "Balance", type: "money", value: (r) => r.Balance },
  {
    key: "transactions",
    label: "Transactions",
    type: "action",
    value: () => "Recent History",
  },
  {
    key: "leaseStart",
    label: "Lease Start",
    type: "date",
    value: (r) => r.TenantName && r.TenantName.leaseStartDate,
  },
  {
    key: "leaseEnd",
    label: "Lease End",
    type: "date",
    value: (r) => r.TenantName && r.TenantName.leaseEndDate,
  },
  {
    key: "moveIn",
    label: "Move-In",
    type: "date",
    value: (r) => r.TenantName && r.TenantName.moveInDate,
  },
  {
    key: "moveOut",
    label: "Move-Out",
    type: "date",
    value: (r) => r.TenantName && r.TenantName.moveOutDate,
  },
  {
    key: "prospectSource",
    label: "Prospect Source",
    type: "text",
    value: (r) => r.TenantName && r.TenantName.prospectSourceName,
  },
  {
    key: "residencyStatus",
    label: "Residency Status",
    type: "text",
    value: residency,
  },
  { key: "dhaStatus", label: "DHA Status", type: "text", value: dha },
  {
    key: "programDocuments",
    label: "Program Documents",
    type: "link",
    value: (r) => r,
  },
  {
    key: "monthlyRent",
    label: "Monthly Rent",
    type: "money",
    value: (r) => r.MonthlyRent,
  },
  {
    key: "hapPortion",
    label: "HAP Portion",
    type: "money",
    value: (r) => r.HAPPortion,
  },
  {
    key: "tenantPortion",
    label: "Tenant Portion",
    type: "money",
    value: (r) => r.HAPAmendment1Portion,
  },
  {
    key: "caseworkerName",
    label: "Caseworker Name",
    type: "text",
    value: (r) => r.CaseworkerName,
  },
  {
    key: "caseworkerContact",
    label: "Caseworker Contact",
    type: "text",
    value: (r) => r.CaseworkerContact,
  },
  {
    key: "hapStart",
    label: "HAP Start",
    type: "date",
    value: (r) => r.HAPContractStart,
  },
  {
    key: "hapEnd",
    label: "HAP End",
    type: "date",
    value: (r) => r.HAPContractEnd,
  },
  {
    key: "applicationFee",
    label: "Application Fee",
    type: "money",
    value: (r) => r.ApplicationFee,
  },
  {
    key: "administrativeFee",
    label: "Administrative Fee",
    type: "money",
    value: (r) => r.AdministrativeFee,
  },
  { key: "riskFee", label: "Risk Fee", type: "money", value: (r) => r.RiskFee },
  {
    key: "securityDeposit",
    label: "Security Deposit",
    type: "money",
    value: (r) => r.SecurityDeposit,
  },
  {
    key: "landlordIncentiveFee",
    label: "Landlord Incentive Fee",
    type: "money",
    value: (r) => r.LandlordIncentiveFee,
  },
  {
    key: "hapAmend1Date",
    label: "HAP Amend 1 Date",
    type: "date",
    value: (r) => r.HAPAmendment1Date,
  },
  {
    key: "hapPortion1",
    label: "HAP Portion 1",
    type: "money",
    value: (r) => r.HAPAmendment1Portion,
  },
  {
    key: "hapAmend2Date",
    label: "HAP Amend 2 Date",
    type: "date",
    value: (r) => r.HAPAmendment2Date,
  },
  {
    key: "hapPortion2",
    label: "HAP Portion 2",
    type: "money",
    value: (r) => r.HAPAmendment2Portion,
  },
  {
    key: "hapAmend3Date",
    label: "HAP Amend 3 Date",
    type: "date",
    value: (r) => r.HAPAmendment3Date,
  },
  {
    key: "hapPortion3",
    label: "HAP Portion 3",
    type: "money",
    value: (r) => r.HAPAmendment3Portion,
  },
];
const initialColumns = col
  .filter(
    (item) =>
      [
        "property",
        "unit",
        "hapAmend1Date",
        "hapPortion1",
        "hapAmend2Date",
        "hapPortion2",
        "hapAmend3Date",
        "hapPortion3",
      ].indexOf(item.key) < 0
  )
  .map((item) => item.key);

export default function Dashboard(
  props: IDashboardProps
): React.ReactElement<IDashboardProps> {
  const base = props.context.pageContext.web.absoluteUrl;
  const [records, setRecords] = React.useState<Intake[]>([]);
  const [people, setPeople] = React.useState<Person[]>([]);
  const [ledgerEntries, setLedgerEntries] = React.useState<LedgerEntry[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [ledgerLoading, setLedgerLoading] = React.useState(true);
  const [error, setError] = React.useState("");
  const [ledgerError, setLedgerError] = React.useState("");
  const [formError, setFormError] = React.useState("");
  const [lastRefresh, setLastRefresh] = React.useState(
    new Date().toLocaleTimeString()
  );
  const [search, setSearch] = React.useState("");
  const [columns, setColumns] = React.useState<string[]>(initialColumns);
  const [headerFilters, setHeaderFilters] = React.useState(false);
  const [columnFilters, setColumnFilters] = React.useState<{
    [key: string]: string;
  }>({});
  const [ledgerHeaderFilters, setLedgerHeaderFilters] = React.useState(false);
  const [ledgerColumnFilters, setLedgerColumnFilters] = React.useState<{
    [key: string]: string;
  }>({});
  const [sort, setSort] = React.useState({ key: "resident", desc: false });
  const [page, setPage] = React.useState(1);
  const [pageSize, setPageSize] = React.useState(pageSizeOptions[0]);
  const [ledgerPage, setLedgerPage] = React.useState(1);
  const [ledgerPageSize, setLedgerPageSize] = React.useState(
    pageSizeOptions[0]
  );
  const [ledgerMonth, setLedgerMonth] = React.useState(currentLedgerMonthKey);
  const [dataView, setDataView] = React.useState<DataView>("resident");
  const [selectedLedgerRow, setSelectedLedgerRow] =
    React.useState<LedgerPivotRow | undefined>();
  const [ledgerDetailLoading, setLedgerDetailLoading] = React.useState(false);
  const [ledgerDetailError, setLedgerDetailError] = React.useState("");
  const [ledgerDetailSort, setLedgerDetailSort] = React.useState<{
    key: LedgerDetailSortKey;
    desc: boolean;
  }>({ key: "date", desc: true });
  const [properties, setProperties] = React.useState<string[]>(() =>
    readArray("dha-properties", [])
  );
  const [residencies, setResidencies] = React.useState<string[]>(() =>
    readArray("dha-residencies", defaultResidency)
  );
  const [statuses, setStatuses] = React.useState<string[]>(() =>
    readArray("dha-statuses", defaultDha)
  );
  const [advanced, setAdvanced] = React.useState(false);
  const [documentStatuses, setDocumentStatuses] = React.useState<string[]>([]);
  const [balanceFilter, setBalanceFilter] = React.useState("all");
  const [leaseFilter, setLeaseFilter] = React.useState("all");
  const [edit, setEdit] = React.useState<Intake | undefined>();
  const [draft, setDraft] = React.useState<{ [key: string]: any }>({});
  const [create, setCreate] = React.useState(false);
  const [tenantSearch, setTenantSearch] = React.useState("");
  const [saving, setSaving] = React.useState(false);
  const [refreshBalancesEndpoint, setRefreshBalancesEndpoint] =
    React.useState("");
  const [llmRequestEndpoint, setLlmRequestEndpoint] = React.useState("");
  const [summarizeTransactionsEndpoint, setSummarizeTransactionsEndpoint] =
    React.useState("");
  const [analyzeBalanceInstructions, setAnalyzeBalanceInstructions] =
    React.useState("");
  const [refreshConfigurationLoading, setRefreshConfigurationLoading] =
    React.useState(true);
  const [refreshConfigurationError, setRefreshConfigurationError] =
    React.useState("");
  const [canManageRefreshConfiguration, setCanManageRefreshConfiguration] =
    React.useState(false);
  const [manageRefreshConfiguration, setManageRefreshConfiguration] =
    React.useState(false);
  const [refreshEndpointDraft, setRefreshEndpointDraft] = React.useState("");
  const [llmRequestEndpointDraft, setLlmRequestEndpointDraft] =
    React.useState("");
  const [summarizeTransactionsEndpointDraft, setSummarizeTransactionsEndpointDraft] =
    React.useState("");
  const [analyzeBalanceInstructionsDraft, setAnalyzeBalanceInstructionsDraft] =
    React.useState("");
  const [refreshEndpointError, setRefreshEndpointError] = React.useState("");
  const [refreshEndpointSaving, setRefreshEndpointSaving] =
    React.useState(false);
  const [refreshBalancesRunning, setRefreshBalancesRunning] =
    React.useState(false);
  const [refreshBalancesStatus, setRefreshBalancesStatus] = React.useState("");
  const [refreshBalancesError, setRefreshBalancesError] = React.useState("");
  const [balanceDetailsLoading, setBalanceDetailsLoading] =
    React.useState(false);
  const [balanceDetailsError, setBalanceDetailsError] = React.useState("");
  const [balanceDetails, setBalanceDetails] = React.useState("");
  const [balanceDetailsVisible, setBalanceDetailsVisible] =
    React.useState(false);
  const [balanceDetailsCopied, setBalanceDetailsCopied] =
    React.useState(false);
  const commandSentinel = React.useRef<HTMLDivElement>(null);
  const columnMenu = React.useRef<HTMLDetailsElement>(null);
  const balanceDetailsContent = React.useRef<HTMLDivElement>(null);
  const ledgerRequestId = React.useRef(0);
  const ledgerDetailRequestId = React.useRef(0);
  const [commandBarPinned, setCommandBarPinned] = React.useState(false);
  React.useEffect(() => {
    store("dha-properties", properties);
  }, [properties]);
  React.useEffect(() => {
    store("dha-residencies", residencies);
  }, [residencies]);
  React.useEffect(() => {
    store("dha-statuses", statuses);
  }, [statuses]);
  React.useEffect(() => {
    const sentinel = commandSentinel.current;
    if (!sentinel || typeof IntersectionObserver === "undefined") return;
    const observer = new IntersectionObserver((entries) =>
      setCommandBarPinned(!entries[0].isIntersecting)
    );
    observer.observe(sentinel);
    return () => observer.disconnect();
  }, []);
  React.useEffect(() => {
    const closeColumnMenu = (event: MouseEvent): void => {
      if (
        columnMenu.current &&
        !columnMenu.current.contains(event.target as Node)
      )
        columnMenu.current.removeAttribute("open");
    };
    document.addEventListener("mousedown", closeColumnMenu);
    return () => document.removeEventListener("mousedown", closeColumnMenu);
  }, []);
  const get = React.useCallback(
    async (url: string): Promise<any> => {
      const response: SPHttpClientResponse =
        await props.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1
        );
      if (!response.ok) {
        const text = await response.text();
        throw new Error(
          `SharePoint returned ${response.status}: ${text.substring(0, 160)}`
        );
      }
      return response.json();
    },
    [props.context.spHttpClient]
  );
  const getAllItems = React.useCallback(
    async (url: string): Promise<any[]> => {
      const items: any[] = [];
      let nextUrl = url;
      while (nextUrl) {
        const result = await get(nextUrl);
        items.push(...(result.value || []));
        nextUrl =
          result["@odata.nextLink"] ||
          result["odata.nextLink"] ||
          (result.d && result.d.__next) ||
          "";
      }
      return items;
    },
    [get]
  );
  const resolveLookupDomainId = React.useCallback(
    async (
      item: Intake | Person,
      listName: string,
      fieldNames: string[]
    ): Promise<string | undefined> => {
      const normalizedFieldNames = fieldNames.map((fieldName) =>
        fieldName.replace(/[^a-z0-9]/gi, "").toLowerCase()
      );
      const fieldsResult = await get(
        `${base}/_api/web/lists/getbytitle('${safe(
          listName
        )}')/fields?$select=InternalName,Title,LookupList,LookupField,TypeAsString`
      );
      const fields = fieldsResult.value ||
        (fieldsResult.d && fieldsResult.d.results) || [];

      for (const field of fields) {
        const fieldNamesToMatch = [field.InternalName, field.Title].map(
          (fieldName) => String(fieldName || "")
            .replace(/[^a-z0-9]/gi, "")
            .toLowerCase()
        );
        if (!fieldNamesToMatch.some((fieldName) =>
          normalizedFieldNames.some(
            (expectedFieldName) =>
              fieldName === expectedFieldName ||
              fieldName.slice(-expectedFieldName.length) === expectedFieldName
          )
        )) continue;

        if (field.TypeAsString !== "Lookup") {
          const fieldValue = domainIdentifier(item[field.InternalName]);
          if (fieldValue) return fieldValue;
          continue;
        }

        const lookupId = item[`${field.InternalName}Id`] ?? item[field.InternalName];
        const directValue = domainIdentifier(lookupId);
        if (directValue) return directValue;
        const numericLookupId = Number(lookupId);
        if (!isFinite(numericLookupId) || !field.LookupList || !field.LookupField)
          continue;

        const lookupListId = String(field.LookupList).replace(/[{}]/g, "");
        const lookupItem = await get(
          `${base}/_api/web/lists(guid'${lookupListId}')/items(${numericLookupId})?$select=${field.LookupField}`
        );
        const resolvedValue = domainIdentifier(lookupItem[field.LookupField]);
        if (resolvedValue) return resolvedValue;
      }
      return undefined;
    },
    [base, get]
  );
  const refreshConfigurationListUrl =
    `${base}/_api/web/lists/getbytitle('${safe(configurationListName)}')`;
  const loadRefreshConfiguration = React.useCallback(async (): Promise<void> => {
    setRefreshConfigurationLoading(true);
    setRefreshConfigurationError("");
    let configurationError = "";
    let isSiteAdmin = false;
    let hasManagePermissions = false;
    let siteAdminLookupFailed = false;
    let permissionsLookupSucceeded = false;
    try {
      const currentUserResult = await get(
        `${base}/_api/web/currentuser?$select=IsSiteAdmin`
      );
      const currentUser = currentUserResult.d || currentUserResult;
      isSiteAdmin = currentUser.IsSiteAdmin === true;
    } catch (e) {
      siteAdminLookupFailed = true;
    }
    const permissionUrls = [
      `${refreshConfigurationListUrl}/EffectiveBasePermissions`,
      `${base}/_api/web/EffectiveBasePermissions`,
    ];
    for (const permissionUrl of permissionUrls) {
      try {
        const permissionResult = await get(permissionUrl);
        const permissionPayload = permissionResult.d || permissionResult;
        const effectivePermissions =
          permissionPayload.EffectiveBasePermissions || permissionPayload;
        const lowPermissions = Number(effectivePermissions.Low || 0);
        permissionsLookupSucceeded = true;
        hasManagePermissions =
          hasManagePermissions ||
          (lowPermissions & managePermissionsMask) === managePermissionsMask;
      } catch (e) {
        // Try the next applicable SharePoint permission scope.
      }
    }
    setCanManageRefreshConfiguration(isSiteAdmin || hasManagePermissions);
    if (siteAdminLookupFailed && !permissionsLookupSucceeded) {
      configurationError =
        "Unable to verify Refresh Balances configuration permissions.";
    }
    try {
      const configurationResult = await get(
        `${refreshConfigurationListUrl}/items?$filter=(Title eq '${safe(
          refreshBalancesConfigurationTitle
        )}' or Title eq '${safe(
          llmRequestEndpointConfigurationTitle
        )}' or Title eq '${safe(
          summarizeTransactionsEndpointConfigurationTitle
        )}' or Title eq '${safe(
          analyzeBalanceInstructionsConfigurationTitle
        )}')`
      );
      const configurationItems = configurationResult.value ||
        (configurationResult.d && configurationResult.d.results) || [];
      const configurationValue = (title: string): string => String(
        (configurationItems.find((item: any) => item.Title === title) || {})[
          configurationValueFieldKey
        ] || ""
      ).trim();
      setRefreshBalancesEndpoint(
        configurationValue(refreshBalancesConfigurationTitle)
      );
      setLlmRequestEndpoint(configurationValue(llmRequestEndpointConfigurationTitle));
      setSummarizeTransactionsEndpoint(
        configurationValue(summarizeTransactionsEndpointConfigurationTitle)
      );
      setAnalyzeBalanceInstructions(
        configurationValue(analyzeBalanceInstructionsConfigurationTitle)
      );
    } catch (e) {
      configurationError = e instanceof Error
        ? e.message
        : "Unable to load the Refresh Balances configuration.";
    } finally {
      setRefreshConfigurationError(configurationError);
      setRefreshConfigurationLoading(false);
    }
  }, [base, get, refreshConfigurationListUrl]);
  React.useEffect(() => {
    void loadRefreshConfiguration();
  }, [loadRefreshConfiguration]);
  const load = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setError("");
    try {
      const [intake, tenant] = await Promise.all([
        get(
          `${base}/_api/web/lists/getbytitle('${safe(
            props.intakeListName || "DHA Intake"
          )}')/items?$top=5000`
        ),
        get(
          `${base}/_api/web/lists/getbytitle('${safe(
            props.peopleListName || "ResMan People"
          )}')/items?$top=5000`
        ),
      ]);
      const peopleData: Person[] = tenant.value || [];
      const lookup: { [key: number]: Person } = {};
      peopleData.forEach((person) => {
        lookup[person.Id] = person;
      });
      setPeople(peopleData);
      setRecords(
        (intake.value || []).map((item: Intake) => ({
          ...item,
          TenantName: lookup[item.TenantNameId || 0],
        }))
      );
      setLastRefresh(new Date().toLocaleTimeString());
    } catch (e) {
      setError(
        e instanceof Error ? e.message : "Unable to load SharePoint data."
      );
    } finally {
      setLoading(false);
    }
  }, [
    base,
    get,
    props.intakeListName,
    props.peopleListName,
  ]);
  const loadLedger = React.useCallback(async (): Promise<void> => {
    const requestId = ++ledgerRequestId.current;
    setLedgerLoading(true);
    setLedgerError("");
    setLedgerEntries([]);
    setSelectedLedgerRow(undefined);
    try {
      const listName = safe(props.ledgerListName || "ResMan TLedger");
      const range = ledgerMonthRange(ledgerMonth);
      const monthResults = await getAllItems(
        `${base}/_api/web/lists/getbytitle('${listName}')/items?$top=5000&$select=Id,Title,TenantNameId,date,transactionType,CategoryName,Description,Amount,PaymentMethod,DateReversed&$filter=date ge datetime'${range.start}' and date lt datetime'${range.end}'&$orderby=date desc`
      );
      if (requestId !== ledgerRequestId.current) return;
      setLedgerEntries(monthResults as LedgerEntry[]);
      setSelectedLedgerRow(undefined);
      setLastRefresh(new Date().toLocaleTimeString());
    } catch (e) {
      if (requestId !== ledgerRequestId.current) return;
      setLedgerEntries([]);
      setLedgerError(
        e instanceof Error ? e.message : "Unable to load tenant ledger data."
      );
    } finally {
      if (requestId === ledgerRequestId.current) setLedgerLoading(false);
    }
  }, [base, getAllItems, ledgerMonth, props.ledgerListName]);
  React.useEffect(() => {
    void load();
  }, [load]);
  React.useEffect(() => {
    void loadLedger();
  }, [loadLedger]);
  const propertyOptions = unique(records.map(prop));
  const residencyOptions = unique(records.map(residency));
  const statusOptions = unique(records.map(dha).concat(["Unknown"]));
  const filtered = records
    .filter(
      (item) =>
        (!properties.length || properties.indexOf(prop(item)) >= 0) &&
        (!residencies.length || residencies.indexOf(residency(item)) >= 0) &&
        (!statuses.length || statuses.indexOf(dha(item)) >= 0)
    )
    .filter(
      (item) =>
        !documentStatuses.length ||
        documentStatuses.every((status) => matchesDocumentStatus(item, status))
    )
    .filter(
      (item) =>
        balanceFilter === "all" ||
        (balanceFilter === "positive"
          ? num(item.Balance) > 0
          : num(item.Balance) < 0)
    )
    .filter((item) => leaseFilter === "all" || leaseMatches(item, leaseFilter))
    .filter(
      (item) =>
        !search.trim() ||
        `${name(item)} ${unit(item)} ${prop(item)}`
          .toLowerCase()
          .indexOf(search.toLowerCase()) >= 0
    )
    .filter((item) =>
      columns.every(
        (key) =>
          !columnFilters[key] ||
          columnFilters[key] === "all" ||
          display(item, getColumn(key), base) === columnFilters[key]
      )
    )
    .sort((a, b) => compare(a, b, sort, base));
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const currentPage = Math.min(page, totalPages);
  const rows = filtered.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize
  );
  const ledgerPivot = buildLedgerPivot(ledgerEntries, people);
  const ledgerSearch = search.trim().toLowerCase();
  const ledgerFilterColumns: Array<{
    key: string;
    value: (row: LedgerPivotRow) => number | undefined;
  }> = [
    {
      key: "total-charges",
      value: (row: LedgerPivotRow) => ledgerValueTotal(row.charges),
    },
    {
      key: "total-payments",
      value: (row: LedgerPivotRow) => ledgerValueTotal(row.payments),
    },
    {
      key: "total-credits",
      value: (row: LedgerPivotRow) => ledgerValueTotal(row.credits),
    },
  ]
    .concat(
      ledgerPivot.chargeCategories.map((category) => ({
        key: `charge-${category}`,
        value: (row: LedgerPivotRow) => row.charges[category],
      }))
    )
    .concat(
      ledgerPivot.paymentCategories.map((category) => ({
        key: `payment-${category}`,
        value: (row: LedgerPivotRow) => row.payments[category],
      }))
    )
    .concat(
      ledgerPivot.creditCategories.map((category) => ({
        key: `credit-${category}`,
        value: (row: LedgerPivotRow) => row.credits[category],
      }))
    );
  const searchedLedgerRows = ledgerPivot.rows.filter((row) => {
    if (!ledgerSearch) return true;
    const transactionText = row.transactions
      .map((transaction) =>
        [
          transaction.transactionType,
          transaction.CategoryName,
          transaction.Description,
          transaction.PaymentMethod,
        ].join(" ")
      )
      .join(" ");
    return [row.tenant, row.unit, row.property, row.month, transactionText]
      .join(" ")
      .toLowerCase()
      .indexOf(ledgerSearch) >= 0;
  });
  const filteredLedgerRows = searchedLedgerRows.filter((row) =>
    ledgerFilterColumns.every((column) => {
      const filter = ledgerColumnFilters[column.key];
      const value = column.value(row);
      return !filter || filter === "all" || String(value) === filter;
    })
  );
  const ledgerTotalPages = Math.max(
    1,
    Math.ceil(filteredLedgerRows.length / ledgerPageSize)
  );
  const currentLedgerPage = Math.min(ledgerPage, ledgerTotalPages);
  const ledgerRows = filteredLedgerRows.slice(
    (currentLedgerPage - 1) * ledgerPageSize,
    currentLedgerPage * ledgerPageSize
  );
  const sortedLedgerTransactions = selectedLedgerRow
    ? selectedLedgerRow.transactions.slice().sort((left, right) =>
        compareLedgerTransactions(left, right, ledgerDetailSort)
      )
    : [];
  const changeLedgerDetailSort = (key: LedgerDetailSortKey): void => {
    setLedgerDetailSort((current) => ({
      key,
      desc: current.key === key ? !current.desc : key === "date",
    }));
  };
  const openLedgerDetails = (row: LedgerPivotRow): void => {
    ++ledgerDetailRequestId.current;
    setLedgerDetailLoading(false);
    setLedgerDetailError("");
    setLedgerDetailSort({ key: "date", desc: true });
    const billingAccountId = Number(row.transactions[0]?.TenantNameId || 0);
    let resident: Intake | undefined;
    records.some((item) => {
      if (Number(item.TenantNameId || 0) !== billingAccountId) return false;
      resident = item;
      return true;
    });
    setSelectedLedgerRow({
      ...row,
      balance: resident ? num(resident.Balance) : undefined,
    });
  };
  const closeLedgerDetails = (): void => {
    ++ledgerDetailRequestId.current;
    setLedgerDetailLoading(false);
    setLedgerDetailError("");
    setSelectedLedgerRow(undefined);
  };
  const openResidentLedgerDetails = async (item: Intake): Promise<void> => {
    const requestId = ++ledgerDetailRequestId.current;
    const billingAccountId = Number(item.TenantNameId || 0);
    const rangeEnd = new Date();
    const rangeStart = new Date(rangeEnd.getTime());
    rangeStart.setUTCFullYear(rangeStart.getUTCFullYear() - 2);
    setLedgerDetailSort({ key: "date", desc: true });
    setLedgerDetailError("");
    setLedgerDetailLoading(true);
    setSelectedLedgerRow({
      key: `resident-${item.Id}`,
      tenant: name(item),
      property: prop(item),
      unit: unit(item),
      balance: num(item.Balance),
      monthKey: "past-two-years",
      month: `${rangeStart.toLocaleDateString("en-US")} to ${rangeEnd.toLocaleDateString("en-US")}`,
      charges: {},
      payments: {},
      credits: {},
      transactions: [],
    });
    if (!billingAccountId) {
      setLedgerDetailLoading(false);
      setLedgerDetailError("This resident does not have a billing account ID.");
      return;
    }
    try {
      const listName = safe(props.ledgerListName || "ResMan TLedger");
      const transactions = await getAllItems(
        `${base}/_api/web/lists/getbytitle('${listName}')/items?$top=5000&$select=Id,Title,TenantNameId,date,transactionType,CategoryName,Description,Amount,PaymentMethod,DateReversed&$filter=TenantNameId eq ${billingAccountId} and date ge datetime'${rangeStart.toISOString()}' and date le datetime'${rangeEnd.toISOString()}'&$orderby=date desc`
      );
      if (requestId !== ledgerDetailRequestId.current) return;
      setSelectedLedgerRow((current) =>
        current ? { ...current, transactions: transactions as LedgerEntry[] } : current
      );
    } catch (e) {
      if (requestId !== ledgerDetailRequestId.current) return;
      setLedgerDetailError(
        e instanceof Error ? e.message : "Unable to load billing account transactions."
      );
    } finally {
      if (requestId === ledgerDetailRequestId.current)
        setLedgerDetailLoading(false);
    }
  };
  const summarizeTransactions = async (item: Intake): Promise<void> => {
    if (balanceDetailsLoading) return;
    setBalanceDetailsVisible(true);
    setBalanceDetails("");
    setBalanceDetailsError("");
    setBalanceDetailsCopied(false);
    setBalanceDetailsLoading(true);
    const userEmail = props.context.pageContext.user.email;
    if (!summarizeTransactionsEndpoint) {
      setBalanceDetailsError(
        "Configure the Summarize Transactions Endpoint before requesting balance details."
      );
      setBalanceDetailsLoading(false);
      return;
    }
    if (!userEmail) {
      setBalanceDetailsError(
        "The selected resident is missing a property ID, billing account ID, or current-user email."
      );
      setBalanceDetailsLoading(false);
      return;
    }
    try {
      const peoplePropertyId = item.TenantName
        ? await resolveLookupDomainId(
            item.TenantName,
            props.peopleListName || "ResMan People",
            ["propertyId"]
          )
        : undefined;
      const intakePropertyId = await resolveLookupDomainId(
        item,
        props.intakeListName || "DHA Intake",
        ["propertyId"]
      );
      const propertyId = [
        peoplePropertyId,
        intakePropertyId,
        residentPropertyId(item),
      ]
        .map(propertyGuid)
        .filter((identifier): identifier is string => Boolean(identifier))[0];
      const billingAccountId =
        (item.TenantName &&
          (await resolveLookupDomainId(
            item.TenantName,
            props.peopleListName || "ResMan People",
            ["billingAccountId", "billingAccount"]
          ))) ||
        (await resolveLookupDomainId(
          item,
          props.intakeListName || "DHA Intake",
          ["billingAccountId", "billingAccount"]
        )) ||
        domainIdentifier(residentBillingAccountId(item));
      if (!propertyId || !billingAccountId) {
        throw new Error(
          "The selected resident is missing a property ID or billing account ID."
        );
      }
      const tokenProvider =
        await props.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken(flowServiceResource);
      const response = await window.fetch(summarizeTransactionsEndpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ propertyId, billingAccountId, userEmail }),
      });
      if (!response.ok) {
        const responseText = await response.text();
        throw new Error(
          `Balance details request failed (${response.status}): ${responseText.substring(
            0,
            160
          )}`
        );
      }
      const responseText = await response.text();
      let summaryResponse: { Status?: unknown; Response?: unknown };
      try {
        summaryResponse = JSON.parse(responseText);
      } catch {
        throw new Error("Balance details returned an invalid response.");
      }
      if (summaryResponse.Status !== "Success") {
        throw new Error(
          `Balance details request was not successful: ${String(
            summaryResponse.Status || "Unknown status"
          )}`
        );
      }
      if (typeof summaryResponse.Response !== "string") {
        throw new Error("Balance details returned no Markdown response.");
      }
      setBalanceDetails(normalizeMarkdown(summaryResponse.Response));
    } catch (e) {
      setBalanceDetailsError(
        e instanceof Error
          ? e.message
          : "Unable to request balance details for this resident."
      );
    } finally {
      setBalanceDetailsLoading(false);
    }
  };
  const copyBalanceDetails = async (): Promise<void> => {
    try {
      const content = balanceDetailsContent.current;
      if (!content) throw new Error("Balance details are not available to copy.");
      const plainText = content.innerText;
      if (
        navigator.clipboard.write &&
        typeof ClipboardItem !== "undefined"
      ) {
        const richText = `<!DOCTYPE html><html><body>${content.innerHTML}</body></html>`;
        await navigator.clipboard.write([
          new ClipboardItem({
            "text/html": new Blob([richText], { type: "text/html" }),
            "text/plain": new Blob([plainText], { type: "text/plain" }),
          }),
        ]);
      } else {
        await navigator.clipboard.writeText(plainText);
      }
      setBalanceDetailsCopied(true);
    } catch {
      setBalanceDetailsCopied(false);
    }
  };
  const activeColumns = col.filter((item) => columns.indexOf(item.key) >= 0);
  const exportCsv = (): void => {
    const headers = ["Resident"].concat(
      activeColumns.map((column) => column.label)
    );
    const dataRows = filtered.map((item) =>
      [name(item)].concat(
        activeColumns.map((column) =>
          column.type === "link"
            ? docUrl(base, item)
            : display(item, column, base)
        )
      )
    );
    const csv = [headers].concat(dataRows).map((row) =>
      row.map(csvCell).join(",")
    ).join("\r\n");
    const blob = new Blob(["\ufeff", csv], {
      type: "text/csv;charset=utf-8",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `dha-resident-data-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => window.URL.revokeObjectURL(url), 0);
  };
  const exportLedgerCsv = (): void => {
    const headers = [
      "Resident",
      "Unit",
      "Property",
      "Month",
      "Total Charges",
      "Total Payments",
      "Total Credits",
      "NET Balance",
    ]
      .concat(
        ledgerPivot.chargeCategories.map((category) => `Charge - ${category}`)
      )
      .concat(
        ledgerPivot.paymentCategories.map(
          (category) => `Payment - ${category}`
        )
      )
      .concat(
        ledgerPivot.creditCategories.map((category) => `Credit - ${category}`)
      );
    const dataRows = filteredLedgerRows.map((row) =>
      [
        row.tenant,
        row.unit,
        row.property,
        row.month,
        String(ledgerValueTotal(row.charges)),
        String(ledgerValueTotal(row.payments)),
        String(ledgerValueTotal(row.credits)),
        String(ledgerNetBalance(row)),
      ]
        .concat(
          ledgerPivot.chargeCategories.map((category) =>
            row.charges[category] === undefined
              ? ""
              : String(row.charges[category])
          )
        )
        .concat(
          ledgerPivot.paymentCategories.map((category) =>
            row.payments[category] === undefined
              ? ""
              : String(row.payments[category])
          )
        )
        .concat(
          ledgerPivot.creditCategories.map((category) =>
            row.credits[category] === undefined
              ? ""
              : String(row.credits[category])
          )
        )
    );
    const csv = [headers]
      .concat(dataRows)
      .map((row) => row.map(csvCell).join(","))
      .join("\r\n");
    const blob = new Blob(["\ufeff", csv], {
      type: "text/csv;charset=utf-8",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `dha-transaction-ledger-${ledgerMonth}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => window.URL.revokeObjectURL(url), 0);
  };
  const exportLedgerDetailsCsv = (): void => {
    if (!selectedLedgerRow || !sortedLedgerTransactions.length) return;
    const headers = [
      "Date",
      "Type",
      "Category Name",
      "Description",
      "Amount",
      "Payment Method",
      "Date Reversed",
    ];
    const dataRows = sortedLedgerTransactions.map((transaction) => [
      date(transaction.date),
      transaction.transactionType || "",
      transaction.CategoryName || "",
      transaction.Description || "",
      String(num(transaction.Amount)),
      transaction.PaymentMethod || "",
      transaction.DateReversed ? date(transaction.DateReversed) : "",
    ]);
    const csv = [headers]
      .concat(dataRows)
      .map((row) => row.map(csvCell).join(","))
      .join("\r\n");
    const blob = new Blob(["\ufeff", csv], {
      type: "text/csv;charset=utf-8",
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `transaction-ledger-details-${new Date()
      .toISOString()
      .slice(0, 10)}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => window.URL.revokeObjectURL(url), 0);
  };
  const kpis = calculateKpis(filtered, base);
  const charts = calculateDashboardCharts(filtered);
  const snapshotMax = Math.max(
    1,
    ...charts.snapshot.map((metric) => metric.value)
  );
  const residentMax = Math.max(
    1,
    ...charts.properties.map((metric) => metric.residents)
  );
  const rentMax = Math.max(
    1,
    ...charts.properties.map((metric) => metric.rent)
  );
  const editChanged =
    Boolean(edit) &&
    (Number(draft.TenantNameId || 0) !== Number(edit!.TenantNameId || 0) ||
      Object.keys(fieldMap).some(
        (key) =>
          comparableFieldValue(fieldMap[key], draft[key]) !==
          comparableFieldValue(fieldMap[key], edit![key])
      ));
  const openEdit = (item: Intake): void => {
    setFormError("");
    setCreate(false);
    setEdit(item);
    setDraft({ ...item });
    setTenantSearch("");
  };
  const openCreate = (): void => {
    setFormError("");
    setCreate(true);
    setEdit(undefined);
    setDraft({ DHAStatus: "Active" });
    setTenantSearch("");
  };
  const openRefreshConfiguration = (): void => {
    setRefreshEndpointDraft(refreshBalancesEndpoint);
    setLlmRequestEndpointDraft(llmRequestEndpoint);
    setSummarizeTransactionsEndpointDraft(summarizeTransactionsEndpoint);
    setAnalyzeBalanceInstructionsDraft(analyzeBalanceInstructions);
    setRefreshEndpointError("");
    setManageRefreshConfiguration(true);
  };
  const saveRefreshConfiguration = async (): Promise<void> => {
    const endpoint = refreshEndpointDraft.trim();
    const llmEndpoint = llmRequestEndpointDraft.trim();
    const summarizeEndpoint = summarizeTransactionsEndpointDraft.trim();
    const instructions = analyzeBalanceInstructionsDraft.trim();
    setRefreshEndpointError("");
    for (const endpointToValidate of [endpoint, llmEndpoint, summarizeEndpoint]) {
      if (!endpointToValidate) continue;
      try {
        const parsedEndpoint = new URL(endpointToValidate);
        if (parsedEndpoint.protocol !== "https:")
          throw new Error("The endpoint must use HTTPS.");
      } catch (e) {
        setRefreshEndpointError(
          e instanceof Error && e.message === "The endpoint must use HTTPS."
            ? e.message
            : "Enter a valid HTTPS endpoint URL."
        );
        return;
      }
    }
    setRefreshEndpointSaving(true);
    try {
      const saveConfiguration = async (title: string, value: string): Promise<void> => {
        const configurationResult = await get(
          `${refreshConfigurationListUrl}/items?$select=Id&$filter=Title eq '${safe(title)}'&$top=1`
        );
        const configurationItems = configurationResult.value ||
          (configurationResult.d && configurationResult.d.results) || [];
        const configurationItem = configurationItems[0];
        const response = await props.context.spHttpClient.fetch(
          configurationItem
            ? `${refreshConfigurationListUrl}/items(${configurationItem.Id})`
            : `${refreshConfigurationListUrl}/items`,
          SPHttpClient.configurations.v1,
          {
            method: "POST",
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-Type": "application/json;odata=nometadata",
              ...(configurationItem
                ? { "IF-MATCH": "*", "X-HTTP-Method": "MERGE" }
                : {}),
            },
            body: JSON.stringify({ Title: title, [configurationValueFieldKey]: value }),
          }
        );
        if (!response.ok) {
          const responseText = await response.text();
          throw new Error(
            `Unable to save ${title} (${response.status}): ${responseText.substring(0, 160)}`
          );
        }
      };
      await Promise.all([
        saveConfiguration(refreshBalancesConfigurationTitle, endpoint),
        saveConfiguration(llmRequestEndpointConfigurationTitle, llmEndpoint),
        saveConfiguration(summarizeTransactionsEndpointConfigurationTitle, summarizeEndpoint),
        saveConfiguration(analyzeBalanceInstructionsConfigurationTitle, instructions),
      ]);
      setRefreshBalancesEndpoint(endpoint);
      setLlmRequestEndpoint(llmEndpoint);
      setSummarizeTransactionsEndpoint(summarizeEndpoint);
      setAnalyzeBalanceInstructions(instructions);
      setRefreshBalancesStatus("");
      setRefreshBalancesError("");
      setManageRefreshConfiguration(false);
    } catch (e) {
      setRefreshEndpointError(
        e instanceof Error
          ? e.message
          : "Unable to save the Refresh Balances endpoint."
      );
    } finally {
      setRefreshEndpointSaving(false);
    }
  };
  const refreshBalances = async (): Promise<void> => {
    if (!refreshBalancesEndpoint) return;
    setRefreshBalancesRunning(true);
    setRefreshBalancesStatus("");
    setRefreshBalancesError("");
    try {
      const tokenProvider =
        await props.context.aadTokenProviderFactory.getTokenProvider();
      const accessToken = await tokenProvider.getToken(flowServiceResource);
      const response = await window.fetch(refreshBalancesEndpoint, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          source: "DHA Portfolio Manager",
          requestedAt: new Date().toISOString(),
          requestedBy: props.context.pageContext.user.email,
        }),
      });
      if (!response.ok) {
        const responseText = await response.text();
        throw new Error(
          `Refresh Balances request failed (${response.status}): ${responseText.substring(
            0,
            160
          )}`
        );
      }
      await load();
      setRefreshBalancesStatus(
        "Balances were refreshed and the Resident Data table was reloaded."
      );
    } catch (e) {
      setRefreshBalancesError(
        e instanceof Error
          ? e.message
          : "Unable to submit the Refresh Balances request."
      );
    } finally {
      setRefreshBalancesRunning(false);
    }
  };
  const save = async (): Promise<void> => {
    setFormError("");
    if (create && !draft.TenantNameId) {
      setFormError("Select a tenant before adding a resident.");
      return;
    }
    const invalid = Object.keys(fieldMap).some(
      (key) =>
        fieldMap[key].type === "money" &&
        draft[key] !== undefined &&
        draft[key] !== "" &&
        !isFinite(Number(draft[key]))
    );
    if (invalid) {
      setFormError("Enter valid numeric amounts before saving.");
      return;
    }
    setSaving(true);
    try {
      const payload: any = {
        Title:
          draft.Title ||
          tenantDisplay(Number(draft.TenantNameId), people) ||
          "New resident",
      };
      Object.keys(fieldMap).forEach((key) => {
        const field = fieldMap[key];
        if (field.type === "money") {
          payload[key] =
            draft[key] === undefined || draft[key] === null || draft[key] === ""
              ? null
              : Number(draft[key]);
        } else if (draft[key] !== undefined && draft[key] !== "") {
          payload[key] = draft[key];
        }
      });
      if (draft.TenantNameId) payload.TenantNameId = Number(draft.TenantNameId);
      const list = safe(props.intakeListName || "DHA Intake");
      const endpoint = create
        ? `${base}/_api/web/lists/getbytitle('${list}')/items`
        : `${base}/_api/web/lists/getbytitle('${list}')/items(${edit!.Id})`;
      const response = await props.context.spHttpClient.fetch(
        endpoint,
        SPHttpClient.configurations.v1,
        {
          method: "POST",
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            ...(create ? {} : { "IF-MATCH": "*", "X-HTTP-Method": "MERGE" }),
          },
          body: JSON.stringify(payload),
        }
      );
      if (!response.ok) throw new Error(`Save failed (${response.status}).`);
      setEdit(undefined);
      setCreate(false);
      await load();
    } catch (e) {
      setFormError(
        e instanceof Error ? e.message : "Unable to save resident record."
      );
    } finally {
      setSaving(false);
    }
  };
  const reset = (): void => {
    setProperties([]);
    setResidencies(defaultResidency);
    setStatuses(defaultDha);
    setDocumentStatuses([]);
    setBalanceFilter("all");
    setLeaseFilter("all");
    setSearch("");
    setLedgerPage(1);
    setColumnFilters({});
    setSort({ key: "resident", desc: false });
    setPage(1);
  };
  return (
    <section className={styles.dashboard}>
      <header className={styles.hero}>
        <div>
          <p className={styles.eyebrow}>EXECUTIVE PORTFOLIO INTELLIGENCE</p>
          <h1>DHA Portfolio Manager</h1>
          <p>
            Executive view of resident assistance coverage, balances, lease
            exposure, and follow-up priorities.
          </p>
        </div>
        <div
          className={`${styles.dataStatus} ${
            loading ? styles.dataStatusLoading : ""
          }`}
          role="status"
          aria-live="polite"
        >
          {loading && (
            <span className={styles.loadingMark} aria-hidden="true">
              <span className={styles.loadingSpinner} />
            </span>
          )}
          <span className={styles.dataStatusCopy}>
            <span>
              {loading ? "Retrieving SharePoint data" : "Data status"}
            </span>
            <strong>
              {loading
                ? "Loading list sources…"
                : `Last refreshed ${lastRefresh}`}
            </strong>
          </span>
        </div>
      </header>
      {error && <div className={styles.error}>{error}</div>}
      {ledgerError && <div className={styles.error}>{ledgerError}</div>}
      <div
        className={styles.commandSentinel}
        ref={commandSentinel}
        aria-hidden="true"
      />
      <section
        className={`${styles.commandBar} ${
          commandBarPinned ? styles.commandBarPinned : ""
        }`}
      >
        <div className={styles.filterSets}>
          <Multi
            label="Properties"
            options={propertyOptions}
            selected={properties}
            onChange={setProperties}
            reset={() => setProperties([])}
          />
          <Multi
            label="Residency"
            options={residencyOptions}
            selected={residencies}
            onChange={setResidencies}
            reset={() => setResidencies(defaultResidency)}
          />
          <Multi
            label="DHA Status"
            options={statusOptions}
            selected={statuses}
            onChange={setStatuses}
            reset={() => setStatuses(defaultDha)}
          />
          <Multi
            label="Documents Status"
            options={documentStatusOptions}
            selected={documentStatuses}
            onChange={setDocumentStatuses}
            reset={() => setDocumentStatuses([])}
          />
        </div>
        <div className={styles.commandActions}>
          <div className={styles.commandSearch}>
            <input
              value={search}
              onChange={(e) => {
                setSearch(e.target.value);
                setPage(1);
                setLedgerPage(1);
              }}
              placeholder="Search residents, units, or transactions…"
              aria-label="Search residents, units, or transactions"
            />
            {search && (
              <button
                type="button"
                className={styles.clearSearch}
                onClick={() => {
                  setSearch("");
                  setPage(1);
                  setLedgerPage(1);
                }}
                aria-label="Clear data search"
                title="Clear search"
              >
                ×
              </button>
            )}
          </div>
          <button
            className={`${styles.refresh} ${
              loading ? styles.refreshLoading : ""
            }`}
            onClick={() => void Promise.all([load(), loadLedger()])}
            disabled={loading || ledgerLoading}
            aria-busy={loading || ledgerLoading}
          >
            <span
              className={`${styles.refreshIcon} ${
                loading || ledgerLoading ? styles.refreshIconLoading : ""
              }`}
              aria-hidden="true"
            >
              ↻
            </span>{" "}
            {loading || ledgerLoading ? "Refreshing…" : "Refresh Data"}
          </button>
          <button onClick={reset}>Reset Filters</button>
          <button onClick={() => setAdvanced(!advanced)}>☷ More Filters</button>
          {canManageRefreshConfiguration && (
            <button onClick={openRefreshConfiguration}>⚙ Manage</button>
          )}
        </div>
        {advanced && (
          <div className={styles.advanced}>
            <Select
              label="Balance Status"
              value={balanceFilter}
              onChange={setBalanceFilter}
              options={[
                ["all", "All"],
                ["positive", "Positive Balance"],
                ["credit", "Credit / Negative Balance"],
              ]}
            />
            <Select
              label="Lease Status"
              value={leaseFilter}
              onChange={setLeaseFilter}
              options={[
                ["all", "All"],
                ["expires90", "Expires Within 90 Days"],
                ["expired", "Expired"],
              ]}
            />
          </div>
        )}
      </section>
      <section className={styles.kpis}>
        {kpis.map((kpi) => (
          <div className={styles.kpi} key={kpi.title}>
            <span>{kpi.title}</span>
            <strong>{kpi.value}</strong>
            <small>{kpi.subtitle}</small>
          </div>
        ))}
      </section>
      <section className={styles.chartGrid}>
        <article className={styles.chartPanel}>
          <header className={styles.chartHeader}>
            <h2>Portfolio snapshot</h2>
          </header>
          <div className={styles.snapshotChart}>
            {charts.snapshot.map((metric) => (
              <div className={styles.snapshotRow} key={metric.label}>
                <div className={styles.chartLabel}>
                  <span>{metric.label}</span>
                  <strong>{metric.value}</strong>
                </div>
                <div className={styles.chartTrack} aria-hidden="true">
                  <span
                    style={{
                      width: `${(metric.value / snapshotMax) * 100}%`,
                      backgroundColor: metric.color,
                    }}
                  />
                </div>
              </div>
            ))}
          </div>
        </article>
        <article className={styles.chartPanel}>
          <header className={styles.chartHeader}>
            <h2>Resident status mix</h2>
          </header>
          <div className={styles.residentMix}>
            {charts.resident.map((group) => {
              const total = group.segments.reduce(
                (sum, segment) => sum + segment.value,
                0
              );
              const description = group.segments
                .map((segment) => `${segment.label}: ${segment.value}`)
                .join(", ");
              return (
                <div className={styles.snapshotRow} key={group.label}>
                  <div className={styles.mixTitle}>{group.label}</div>
                  <div
                    className={styles.mixTrack}
                    role="img"
                    aria-label={`${group.label}. ${description}`}
                  >
                    {group.segments.map((segment) => (
                      <span
                        key={segment.label}
                        title={`${segment.label}: ${segment.value}`}
                        style={{
                          width: `${
                            total ? (segment.value / total) * 100 : 0
                          }%`,
                          backgroundColor: segment.color,
                        }}
                      />
                    ))}
                  </div>
                  <div className={styles.mixLegend}>
                    {group.segments.map((segment) => (
                      <span key={segment.label}>
                        <i style={{ backgroundColor: segment.color }} />
                        <span title={segment.label}>{segment.label}</span>
                        <strong>{segment.value}</strong>
                      </span>
                    ))}
                  </div>
                </div>
              );
            })}
          </div>
        </article>
        <article className={styles.chartPanel}>
          <header className={styles.chartHeader}>
            <h2>Property distribution</h2>
            <span>{charts.properties.length} properties</span>
          </header>
          {charts.properties.length ? (
            <div className={styles.propertyCharts}>
              <div className={styles.propertyChart}>
                <h3>Residents by property</h3>
                {charts.properties.map((metric) => (
                  <div className={styles.propertyRow} key={metric.name}>
                    <span title={metric.name}>{metric.name}</span>
                    <div className={styles.chartTrack} aria-hidden="true">
                      <i
                        style={{
                          width: `${(metric.residents / residentMax) * 100}%`,
                        }}
                      />
                    </div>
                    <strong>{metric.residents}</strong>
                  </div>
                ))}
              </div>
              <div className={styles.propertyChart}>
                <h3>Monthly rent exposure by property</h3>
                {charts.properties.map((metric) => (
                  <div className={styles.propertyRow} key={metric.name}>
                    <span title={metric.name}>{metric.name}</span>
                    <div className={styles.chartTrack} aria-hidden="true">
                      <i
                        style={{ width: `${(metric.rent / rentMax) * 100}%` }}
                      />
                    </div>
                    <strong>{money.format(metric.rent)}</strong>
                  </div>
                ))}
              </div>
            </div>
          ) : (
            <p className={styles.chartEmpty}>No records match the filters.</p>
          )}
        </article>
      </section>
      <nav className={styles.dataViewSwitcher} aria-label="Select data view">
        <div
          role="tablist"
          aria-label="Portfolio data views"
          data-view={dataView}
        >
          <button
            type="button"
            role="tab"
            aria-selected={dataView === "resident"}
            className={dataView === "resident" ? styles.activeDataView : ""}
            onClick={() => setDataView("resident")}
          >
            <span className={styles.viewIcon} aria-hidden="true">⌂</span>
            <span>Resident Data</span>
            <b>{filtered.length}</b>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={dataView === "ledger"}
            className={dataView === "ledger" ? styles.activeDataView : ""}
            onClick={() => setDataView("ledger")}
          >
            <span className={styles.viewIcon} aria-hidden="true">▤</span>
            <span>Transaction Ledger</span>
            <b>{filteredLedgerRows.length}</b>
          </button>
        </div>
        <span className={styles.switcherLabel}>Workspace view</span>
      </nav>
      <div key={dataView} className={styles.tableView}>
      {dataView === "resident" ? (
      <section className={styles.tableCard}>
        <div className={styles.tableHeader}>
          <div>
            <h2>
              Resident Data <span>({filtered.length})</span>
            </h2>
            <p>
              {loading
                ? "Refreshing SharePoint data and lookup values…"
                : "Data is current from loaded SharePoint collections."}
            </p>
          </div>
          <div className={styles.tableActions}>
            <button className={styles.add} onClick={openCreate}>
              ＋ Add resident
            </button>
            <button
              className={`${styles.refresh} ${
                refreshBalancesRunning ? styles.refreshLoading : ""
              }`}
              onClick={() => void refreshBalances()}
              disabled={
                refreshConfigurationLoading ||
                refreshBalancesRunning ||
                !refreshBalancesEndpoint
              }
              aria-busy={refreshBalancesRunning}
            >
              <span
                className={`${styles.refreshIcon} ${
                  refreshBalancesRunning ? styles.refreshIconLoading : ""
                }`}
                aria-hidden="true"
              >
                ↻
              </span>{" "}
              {refreshBalancesRunning ? "Refreshing Balances…" : "Refresh Balances"}
            </button>
            <details className={styles.columnMenu} ref={columnMenu}>
              <summary>▦ Selected Columns</summary>
              <div className={styles.columnPanel}>
                {col.map((item) => (
                  <label key={item.key}>
                    <input
                      type="checkbox"
                      checked={columns.indexOf(item.key) >= 0}
                      onChange={() => {
                        setColumns((current) =>
                          current.indexOf(item.key) >= 0
                            ? current.filter((key) => key !== item.key)
                            : current.concat(item.key)
                        );
                        setPage(1);
                      }}
                    />
                    <span>{item.label}</span>
                  </label>
                ))}
              </div>
            </details>
            <button onClick={() => setHeaderFilters(!headerFilters)}>
              {headerFilters ? "Hide Header Filters" : "Show Header Filters"}
            </button>
            <button onClick={exportCsv} disabled={loading || !filtered.length}>
              ⇩ Export CSV
            </button>
          </div>
        </div>
        {!refreshConfigurationLoading && !refreshBalancesEndpoint && (
          <div className={styles.formError} role="note">
            The endpoint for the Refresh Balances function needs to be
            configured by a site owner.
          </div>
        )}
        {refreshConfigurationError && (
          <div className={styles.formError} role="alert">
            {refreshConfigurationError}
          </div>
        )}
        {refreshBalancesError && (
          <div className={styles.formError} role="alert">
            {refreshBalancesError}
          </div>
        )}
        {refreshBalancesStatus && <p role="status">{refreshBalancesStatus}</p>}
        <div className={styles.scroll}>
          <table>
            <thead>
              <tr>
                <th className={styles.stickyColumn}>
                  <button onClick={() => setSort(nextSort(sort, "resident"))}>
                    Resident{" "}
                    {sort.key === "resident" ? (sort.desc ? "↓" : "↑") : "↕"}
                  </button>
                </th>
                {activeColumns.map((item) => (
                  <th key={item.key}>
                    <button
                      className={item.type === "money" ? styles.pivotValue : undefined}
                      onClick={() => setSort(nextSort(sort, item.key))}
                    >
                      {item.label}{" "}
                      {sort.key === item.key ? (sort.desc ? "↓" : "↑") : "↕"}
                    </button>
                  </th>
                ))}
              </tr>
              {headerFilters && (
                <tr className={styles.headerFilterRow}>
                  <th className={styles.stickyColumn} />
                  {activeColumns.map((item) => (
                    <th key={item.key}>
                      <select
                        value={columnFilters[item.key] || "all"}
                        onChange={(e) => {
                          setColumnFilters((current) => ({
                            ...current,
                            [item.key]: e.target.value,
                          }));
                          setPage(1);
                        }}
                      >
                        <option value="all">All</option>
                        {unique(
                          filtered
                            .map((record) => display(record, item, base))
                            .filter((text) => text !== "—")
                        ).map((text) => (
                          <option key={text} value={text}>
                            {text}
                          </option>
                        ))}
                      </select>
                    </th>
                  ))}
                </tr>
              )}
            </thead>
            <tbody>
              {rows.map((item) => (
                <tr key={item.Id} onClick={() => openEdit(item)}>
                  <td className={styles.stickyColumn}>
                    <strong>{name(item)}</strong>
                    <small>
                      {unit(item) || "Unit"} · {prop(item) || "Property"}
                    </small>
                  </td>
                  {activeColumns.map((column) => {
                    const dateWarning = dateWarningMessage(item, column.key);
                    return (
                      <td
                        key={column.key}
                        className={[
                          column.type === "money" ? styles.pivotValue : "",
                          dateWarning ? styles.dateWarning : "",
                        ]
                          .filter(Boolean)
                          .join(" ") || undefined}
                        title={dateWarning}
                      >
                        {column.type === "link" ? (
                          <a
                            href={docUrl(base, item)}
                            target="_blank"
                            rel="noreferrer"
                            data-interception="off"
                            onClick={(e) => e.stopPropagation()}
                          >
                            View Document
                          </a>
                        ) : column.type === "action" ? (
                          <a
                            href="#"
                            onClick={(event) => {
                              event.preventDefault();
                              event.stopPropagation();
                              void openResidentLedgerDetails(item);
                            }}
                          >
                            Recent History
                          </a>
                        ) : column.key === "balance" && num(column.value(item)) > 0 ? (
                          <a
                            href="#"
                            aria-busy={balanceDetailsLoading}
                            onClick={(event) => {
                              event.preventDefault();
                              event.stopPropagation();
                              void summarizeTransactions(item);
                            }}
                          >
                            {display(item, column, base)}
                          </a>
                        ) : (
                          display(item, column, base)
                        )}
                      </td>
                    );
                  })}
                </tr>
              ))}
              {!loading && !rows.length && (
                <tr>
                  <td
                    colSpan={activeColumns.length + 1}
                    className={styles.empty}
                  >
                    No matching residents found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        <footer>
          <div className={styles.pageSummary}>
            <span>
              Showing {filtered.length ? (currentPage - 1) * pageSize + 1 : 0}–
              {Math.min(currentPage * pageSize, filtered.length)} of{" "}
              {filtered.length} matching residents
            </span>
            <label className={styles.pageSize}>
              Rows per page
              <select
                value={pageSize}
                onChange={(event) => {
                  setPageSize(Number(event.target.value));
                  setPage(1);
                }}
              >
                {pageSizeOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </label>
          </div>
          <div>
            <button disabled={currentPage === 1} onClick={() => setPage(1)}>
              First
            </button>
            <button
              disabled={currentPage === 1}
              onClick={() => setPage(currentPage - 1)}
            >
              Previous
            </button>
            <b>
              Page {currentPage} of {totalPages}
            </b>
            <button
              disabled={currentPage === totalPages}
              onClick={() => setPage(currentPage + 1)}
            >
              Next
            </button>
            <button
              disabled={currentPage === totalPages}
              onClick={() => setPage(totalPages)}
            >
              Last
            </button>
          </div>
        </footer>
      </section>
      ) : (
      <section className={styles.tableCard}>
        <div className={styles.tableHeader}>
          <div>
            <h2>
              Transaction Ledger for {monthLabel(ledgerMonth)} <span>({filteredLedgerRows.length})</span>
            </h2>
            <p>
              {ledgerEntries.length} transactions grouped by tenant, property,
              unit, and month.
            </p>
          </div>
          <div className={styles.tableActions}>
            <MonthSelect
              selected={ledgerMonth}
              onChange={(month) => {
                setLedgerMonth(month);
                setLedgerPage(1);
              }}
            />
            <button onClick={() => setLedgerHeaderFilters(!ledgerHeaderFilters)}>
              {ledgerHeaderFilters
                ? "Hide Header Filters"
                : "Show Header Filters"}
            </button>
            <button
              onClick={exportLedgerCsv}
              disabled={ledgerLoading || !filteredLedgerRows.length}
            >
              ⇩ Export CSV
            </button>
          </div>
        </div>
        <div className={styles.scroll}>
          <table className={styles.ledgerTable}>
            <thead>
              <tr className={styles.pivotGroupHeader}>
                <th
                  rowSpan={ledgerHeaderFilters ? 3 : 2}
                  className={styles.stickyLedgerHeader}
                >
                  <span>Resident</span>
                </th>
                <th rowSpan={ledgerHeaderFilters ? 3 : 2}><span>Month</span></th>
                <th colSpan={3} className={styles.totalsGroup}>
                  <span>Totals</span>
                </th>
                {ledgerPivot.chargeCategories.length > 0 && (
                  <th
                    colSpan={ledgerPivot.chargeCategories.length}
                    className={styles.chargeGroup}
                  >
                    <span>Charges</span>
                  </th>
                )}
                {ledgerPivot.paymentCategories.length > 0 && (
                  <th
                    colSpan={ledgerPivot.paymentCategories.length}
                    className={styles.paymentGroup}
                  >
                    <span>Payments</span>
                  </th>
                )}
                {ledgerPivot.creditCategories.length > 0 && (
                  <th
                    colSpan={ledgerPivot.creditCategories.length}
                    className={styles.creditGroup}
                  >
                    <span>Credits</span>
                  </th>
                )}
              </tr>
              <tr className={styles.pivotCategoryHeader}>
                <th className={styles.totalCategory}><span>Charges</span></th>
                <th className={styles.totalCategory}><span>Payments</span></th>
                <th className={styles.totalCategory}><span>Credits</span></th>
                {ledgerPivot.chargeCategories.map((category) => (
                  <th key={`charge-${category}`}><span>{category}</span></th>
                ))}
                {ledgerPivot.paymentCategories.map((category) => (
                  <th key={`payment-${category}`}><span>{category}</span></th>
                ))}
                {ledgerPivot.creditCategories.map((category) => (
                  <th key={`credit-${category}`}><span>{category}</span></th>
                ))}
              </tr>
              {ledgerHeaderFilters && (
                <tr className={styles.headerFilterRow}>
                  {ledgerFilterColumns.map((column) => (
                    <th key={column.key}>
                      <select
                        aria-label={`Filter ${column.key}`}
                        value={ledgerColumnFilters[column.key] || "all"}
                        onChange={(event) => {
                          setLedgerColumnFilters((current) => ({
                            ...current,
                            [column.key]: event.target.value,
                          }));
                          setLedgerPage(1);
                        }}
                      >
                        <option value="all">All</option>
                        {unique(
                          searchedLedgerRows
                            .map((row) => column.value(row))
                            .filter(
                              (value): value is number => value !== undefined
                            )
                            .map(String)
                        ).map((value) => (
                          <option key={value} value={value}>
                            {money.format(Number(value))}
                          </option>
                        ))}
                      </select>
                    </th>
                  ))}
                </tr>
              )}
            </thead>
            <tbody>
              {ledgerLoading && (
                <tr className={styles.ledgerLoadingRow}>
                  <td
                    colSpan={
                      5 +
                      ledgerPivot.chargeCategories.length +
                      ledgerPivot.paymentCategories.length +
                      ledgerPivot.creditCategories.length
                    }
                  >
                    <div className={styles.ledgerLoadingState} role="status">
                      <span className={styles.ledgerLoadingSpinner} aria-hidden="true" />
                      <span>
                        <strong>Loading transaction ledger</strong>
                        <small>Fetching the selected month from SharePoint…</small>
                      </span>
                    </div>
                  </td>
                </tr>
              )}
              {ledgerRows.map((row) => (
                <tr
                  key={row.key}
                  tabIndex={0}
                  onClick={() => openLedgerDetails(row)}
                  onKeyDown={(event) => {
                    if (event.key === "Enter" || event.key === " ") {
                      event.preventDefault();
                      openLedgerDetails(row);
                    }
                  }}
                  aria-label={`View ${row.transactions.length} transactions for ${row.tenant}, ${row.month}`}
                >
                  <td className={styles.stickyColumn}>
                    <strong>{row.tenant}</strong>
                    <small>
                      {row.unit || "Unit"} · {row.property || "Property"}
                    </small>
                  </td>
                  <td>{row.month}</td>
                  <td className={`${styles.pivotValue} ${styles.totalValue}`}>
                    {money.format(ledgerValueTotal(row.charges))}
                  </td>
                  <td className={`${styles.pivotValue} ${styles.totalValue}`}>
                    {money.format(ledgerValueTotal(row.payments))}
                  </td>
                  <td className={`${styles.pivotValue} ${styles.totalValue}`}>
                    {money.format(ledgerValueTotal(row.credits))}
                  </td>
                  {ledgerPivot.chargeCategories.map((category) => (
                    <td key={`charge-${category}`} className={styles.pivotValue}>
                      {row.charges[category]
                        ? money.format(row.charges[category])
                        : "—"}
                    </td>
                  ))}
                  {ledgerPivot.paymentCategories.map((category) => (
                    <td key={`payment-${category}`} className={styles.pivotValue}>
                      {row.payments[category]
                        ? money.format(row.payments[category])
                        : "—"}
                    </td>
                  ))}
                  {ledgerPivot.creditCategories.map((category) => (
                    <td key={`credit-${category}`} className={styles.pivotValue}>
                      {row.credits[category]
                        ? money.format(row.credits[category])
                        : "—"}
                    </td>
                  ))}
                </tr>
              ))}
              {!ledgerLoading && !ledgerRows.length && (
                <tr>
                  <td
                    colSpan={
                      5 +
                      ledgerPivot.chargeCategories.length +
                      ledgerPivot.paymentCategories.length +
                      ledgerPivot.creditCategories.length
                    }
                    className={styles.empty}
                  >
                    No matching transaction ledger groups found.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        <footer>
          <div className={styles.pageSummary}>
            <span>
              Showing {filteredLedgerRows.length ? (currentLedgerPage - 1) * ledgerPageSize + 1 : 0}–
              {Math.min(currentLedgerPage * ledgerPageSize, filteredLedgerRows.length)} of {filteredLedgerRows.length} groups
            </span>
            <label className={styles.pageSize}>
              Rows per page
              <select
                value={ledgerPageSize}
                onChange={(event) => {
                  setLedgerPageSize(Number(event.target.value));
                  setLedgerPage(1);
                }}
              >
                {pageSizeOptions.map((option) => (
                  <option key={option} value={option}>{option}</option>
                ))}
              </select>
            </label>
          </div>
          <div>
            <button disabled={currentLedgerPage === 1} onClick={() => setLedgerPage(currentLedgerPage - 1)}>
              Previous
            </button>
            <b>Page {currentLedgerPage} of {ledgerTotalPages}</b>
            <button disabled={currentLedgerPage === ledgerTotalPages} onClick={() => setLedgerPage(currentLedgerPage + 1)}>
              Next
            </button>
          </div>
        </footer>
      </section>
      )}
      </div>
      {selectedLedgerRow && (
        <div className={styles.modalBackdrop} role="dialog" aria-modal="true">
          <div className={`${styles.modal} ${styles.ledgerDetailModal}`}>
            <header>
              <div>
                <p className={styles.eyebrow}>TRANSACTION LEDGER DETAILS</p>
                <h2>{selectedLedgerRow.tenant}</h2>
                <small>
                  {selectedLedgerRow.property || "Property"} · {selectedLedgerRow.unit || "Unit"} · {selectedLedgerRow.month}
                </small>
                {selectedLedgerRow.balance !== undefined && (
                  <p className={styles.ledgerDetailBalance}>
                    Current Balance: <strong>{money.format(selectedLedgerRow.balance)}</strong>
                  </p>
                )}
              </div>
              <button
                onClick={closeLedgerDetails}
                aria-label="Close transaction ledger details"
              >
                ×
              </button>
            </header>
            <p className={styles.ledgerDetailSummary}>
              {ledgerDetailLoading
                ? "Loading billing account transactions..."
                : `${selectedLedgerRow.transactions.length} underlying transactions`}
            </p>
            {ledgerDetailError && (
              <div className={styles.error} role="alert">
                {ledgerDetailError}
              </div>
            )}
            {ledgerDetailLoading ? (
              <div className={styles.ledgerLoadingState}>
                <span className={styles.ledgerLoadingSpinner} aria-hidden="true" />
                <span>
                  <strong>Loading transactions</strong>
                  <small>Retrieving activity for this billing account.</small>
                </span>
              </div>
            ) : !ledgerDetailError && (
              <div className={styles.scroll}>
              <table className={styles.ledgerDetailTable}>
                <thead>
                  <tr>
                    {ledgerDetailColumn("date", "Date", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("transactionType", "Type", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("CategoryName", "Category Name", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("Description", "Description", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("Amount", "Amount", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("PaymentMethod", "Payment Method", ledgerDetailSort, changeLedgerDetailSort)}
                    {ledgerDetailColumn("DateReversed", "Date Reversed", ledgerDetailSort, changeLedgerDetailSort)}
                  </tr>
                </thead>
                <tbody>
                  {sortedLedgerTransactions.map((transaction) => (
                      <tr key={transaction.Id}>
                        <td>{date(transaction.date)}</td>
                        <td className={styles.ledgerDetailType}>{transaction.transactionType || "—"}</td>
                        <td>{transaction.CategoryName || "—"}</td>
                        <td className={styles.ledgerDescription}>
                          {transaction.Description || "—"}
                        </td>
                        <td className={styles.pivotValue}>
                          {money.format(num(transaction.Amount))}
                        </td>
                        <td>{transaction.PaymentMethod || "—"}</td>
                        <td>{date(transaction.DateReversed)}</td>
                      </tr>
                    ))}
                  {!sortedLedgerTransactions.length && (
                    <tr>
                      <td colSpan={7} className={styles.empty}>
                        No transactions found for this period.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
              </div>
            )}
            <footer>
              <button
                onClick={exportLedgerDetailsCsv}
                disabled={
                  ledgerDetailLoading ||
                  Boolean(ledgerDetailError) ||
                  !sortedLedgerTransactions.length
                }
              >
                ⇩ Export CSV
              </button>
              <button onClick={closeLedgerDetails}>
                Close
              </button>
            </footer>
          </div>
        </div>
      )}
      {balanceDetailsVisible && (
        <div
          className={styles.modalBackdrop}
          role="dialog"
          aria-modal="true"
          aria-label="Balance details"
        >
          <div className={`${styles.modal} ${styles.analysisModal}`}>
            <header>
              <div>
                <p className={styles.eyebrow}>RESIDENT BALANCE</p>
                <h2>Balance Details</h2>
              </div>
              <button
                onClick={() => setBalanceDetailsVisible(false)}
                aria-label="Close balance details"
              >
                ×
              </button>
            </header>
            {balanceDetailsLoading ? (
              <div className={styles.analysisProgress} role="status">
                <span className={styles.analysisSpinner} aria-hidden="true" />
                <span>Analyzing transactions with AI to generate a summary. This may take a few seconds.</span>
              </div>
            ) : balanceDetailsError ? (
              <div className={styles.error} role="alert">
                {balanceDetailsError}
              </div>
            ) : (
              <div ref={balanceDetailsContent} className={styles.analysisContent}>
                {renderMarkdown(balanceDetails)}
              </div>
            )}
            <footer>
              {!balanceDetailsLoading && !balanceDetailsError && (
                <button
                  onClick={() => void copyBalanceDetails()}
                  disabled={!balanceDetails}
                >
                  {balanceDetailsCopied ? "Copied" : "Copy summary"}
                </button>
              )}
              <button onClick={() => setBalanceDetailsVisible(false)}>
                Close
              </button>
            </footer>
          </div>
        </div>
      )}
      {manageRefreshConfiguration && (
        <div className={styles.modalBackdrop} role="dialog" aria-modal="true">
          <div className={`${styles.modal} ${styles.configurationModal}`}>
            <header>
              <div>
                <p className={styles.eyebrow}>SITE OWNER CONFIGURATION</p>
                <h2>Manage Refresh Balances</h2>
                <small>
                  Configure the Entra-authenticated Power Automate HTTP
                  endpoint for this Resident Data list.
                </small>
              </div>
              <button
                onClick={() => setManageRefreshConfiguration(false)}
                aria-label="Close Refresh Balances configuration"
              >
                ×
              </button>
            </header>
            <div className={`${styles.form} ${styles.configurationForm}`}>
              <label>
                <span>Refresh Balances Endpoint</span>
                <small>Invoked when resident balances are refreshed.</small>
                <input
                  type="url"
                  value={refreshEndpointDraft}
                  onChange={(event) =>
                    setRefreshEndpointDraft(event.target.value)
                  }
                  placeholder="https://…"
                  autoFocus
                />
              </label>
              <label>
                <span>LLM Request Endpoint</span>
                <small>Endpoint for requesting AI capabilities.</small>
                <input
                  type="url"
                  value={llmRequestEndpointDraft}
                  onChange={(event) =>
                    setLlmRequestEndpointDraft(event.target.value)
                  }
                  placeholder="https://…"
                />
              </label>
              <label>
                <span>Summarize Transactions Endpoint</span>
                <small>Used to request a summary of transaction history.</small>
                <input
                  type="url"
                  value={summarizeTransactionsEndpointDraft}
                  onChange={(event) =>
                    setSummarizeTransactionsEndpointDraft(event.target.value)
                  }
                  placeholder="https://…"
                />
              </label>
              <label>
                <span>Analyze Balance Instructions</span>
                <small>System guidance sent with each balance analysis request.</small>
                <textarea
                  value={analyzeBalanceInstructionsDraft}
                  onChange={(event) =>
                    setAnalyzeBalanceInstructionsDraft(event.target.value)
                  }
                  rows={6}
                />
              </label>
            </div>
            {refreshEndpointError && (
              <div className={styles.formError} role="alert">
                {refreshEndpointError}
              </div>
            )}
            <footer>
              <button
                onClick={() => setManageRefreshConfiguration(false)}
                disabled={refreshEndpointSaving}
              >
                Cancel
              </button>
              <button
                className={styles.add}
                onClick={() => void saveRefreshConfiguration()}
                disabled={refreshEndpointSaving}
              >
                {refreshEndpointSaving ? "Saving…" : "Save configuration"}
              </button>
            </footer>
          </div>
        </div>
      )}
      {(edit || create) && (
        <div className={styles.modalBackdrop} role="dialog" aria-modal="true">
          <div className={styles.modal}>
            <header>
              <div>
                <p className={styles.eyebrow}>
                  {create ? "NEW RECORD" : "EDIT RESIDENT"}
                </p>
                <h2>{create ? "Add resident record" : name(edit!)}</h2>
                <small>
                  {!create &&
                    `${unit(edit!) || "Unit"} · ${prop(edit!) || "Property"}`}
                </small>
              </div>
              <button
                onClick={() => {
                  setFormError("");
                  setEdit(undefined);
                  setCreate(false);
                }}
              >
                ×
              </button>
            </header>
            <div className={styles.form}>
              {create && (
                <label>
                  Tenant Name
                  <input
                    list="tenant-list"
                    value={
                      tenantDisplay(Number(draft.TenantNameId), people) ||
                      tenantSearch
                    }
                    onChange={(event) => {
                      setTenantSearch(event.target.value);
                      const found = people.filter(
                        (person) =>
                          tenantLabel(person).toLowerCase() ===
                          event.target.value.toLowerCase()
                      )[0];
                      setDraft((current: any) => ({
                        ...current,
                        TenantNameId: found ? found.Id : undefined,
                      }));
                    }}
                    placeholder="Search tenant name…"
                  />
                  <datalist id="tenant-list">
                    {people
                      .filter(
                        (person) =>
                          !tenantSearch ||
                          tenantLabel(person)
                            .toLowerCase()
                            .indexOf(tenantSearch.toLowerCase()) >= 0
                      )
                      .sort((left, right) =>
                        tenantLabel(left).localeCompare(
                          tenantLabel(right),
                          undefined,
                          { numeric: true, sensitivity: "base" }
                        )
                      )
                      .slice(0, 40)
                      .map((person) => (
                        <option key={person.Id} value={tenantLabel(person)} />
                      ))}
                  </datalist>
                </label>
              )}
              {fieldGroups.map((group) => (
                <section key={group.title}>
                  <h3>{group.title}</h3>
                  {group.fields.map((field) => (
                    <Field
                      key={field.key}
                      field={field}
                      value={draft[field.key]}
                      setValue={(value) =>
                        setDraft((current: any) => ({
                          ...current,
                          [field.key]: value,
                        }))
                      }
                    />
                  ))}
                </section>
              ))}
            </div>
            {formError && (
              <div className={styles.formError} role="alert">
                {formError}
              </div>
            )}
            <footer>
              <button
                onClick={() => {
                  setFormError("");
                  setEdit(undefined);
                  setCreate(false);
                }}
              >
                Cancel
              </button>
              <button
                className={styles.add}
                onClick={() => void save()}
                disabled={saving || (!create && !editChanged)}
              >
                {saving ? "Saving…" : create ? "Add resident" : "Save changes"}
              </button>
            </footer>
          </div>
        </div>
      )}
    </section>
  );
}
function Multi(props: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (values: string[]) => void;
  reset: () => void;
}): React.ReactElement<any> {
  const [open, setOpen] = React.useState(false);
  const [search, setSearch] = React.useState("");
  const container = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (!open) return;
    const closeOutside = (event: MouseEvent): void => {
      if (
        container.current &&
        !container.current.contains(event.target as Node)
      )
        setOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent): void => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", closeOutside);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("mousedown", closeOutside);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [open]);
  const label =
    props.selected.length === 0
      ? `${props.label}: All`
      : props.selected.length <= 2
      ? `${props.label}: ${props.selected.join(", ")}`
      : `${props.label}: ${props.selected.length} selected`;
  const shown = props.options.filter(
    (option) => option.toLowerCase().indexOf(search.toLowerCase()) >= 0
  );
  return (
    <div className={styles.multi} ref={container}>
      <button aria-expanded={open} onClick={() => setOpen(!open)}>
        ⌕ {label}
      </button>
      {open && (
        <div className={styles.multiPanel}>
          <input
            placeholder={`Search ${props.label.toLowerCase()}…`}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
          {shown.map((option) => (
            <label key={option}>
              <input
                type="checkbox"
                checked={props.selected.indexOf(option) >= 0}
                onChange={() =>
                  props.onChange(
                    props.selected.indexOf(option) >= 0
                      ? props.selected.filter((value) => value !== option)
                      : props.selected.concat(option)
                  )
                }
              />{" "}
              {option}
            </label>
          ))}
          <div>
            <button onClick={() => props.onChange([])}>Clear</button>
            <button onClick={props.reset}>Reset default</button>
          </div>
        </div>
      )}
    </div>
  );
}
function MonthSelect(props: {
  selected: string;
  onChange: (value: string) => void;
}): React.ReactElement<any> {
  const [open, setOpen] = React.useState(false);
  const [search, setSearch] = React.useState("");
  const container = React.useRef<HTMLDivElement>(null);
  const currentMonth = currentLedgerMonthKey();
  const options = ledgerMonthOptions();
  React.useEffect(() => {
    if (!open) return;
    const closeOutside = (event: MouseEvent): void => {
      if (
        container.current &&
        !container.current.contains(event.target as Node)
      )
        setOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent): void => {
      if (event.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", closeOutside);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("mousedown", closeOutside);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, [open]);
  const shown = options.filter(
    (option) =>
      option.label.toLowerCase().indexOf(search.toLowerCase()) >= 0 ||
      option.key.indexOf(search) >= 0
  );
  const label = monthLabel(props.selected);
  return (
    <div className={`${styles.multi} ${styles.monthMulti}`} ref={container}>
      <button type="button" aria-expanded={open} onClick={() => setOpen(!open)}>
        ◫ Month: {label}
      </button>
      {open && (
        <div className={styles.multiPanel}>
          <input
            placeholder="Search month or year…"
            value={search}
            onChange={(event) => setSearch(event.target.value)}
          />
          {shown.map((option) => (
            <label key={option.key}>
              <input
                type="radio"
                name="ledger-month"
                checked={props.selected === option.key}
                onChange={() => {
                  props.onChange(option.key);
                  setOpen(false);
                }}
              />
              {option.label}
            </label>
          ))}
          <div>
            <button type="button" onClick={() => {
              props.onChange(currentMonth);
              setOpen(false);
            }}>
              Current month
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
function Select(props: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: string[][];
}): React.ReactElement<any> {
  return (
    <label>
      {props.label}
      <select
        value={props.value}
        onChange={(e) => props.onChange(e.target.value)}
      >
        {props.options.map((option) => (
          <option key={option[0]} value={option[0]}>
            {option[1]}
          </option>
        ))}
      </select>
    </label>
  );
}
function Field(props: {
  field: FormField;
  value: any;
  setValue: (value: any) => void;
}): React.ReactElement<any> {
  if (props.field.type === "status")
    return (
      <Select
        label={props.field.label}
        value={props.value || ""}
        onChange={props.setValue}
        options={[
          ["", "Blank"],
          ["Active", "Active"],
          ["Expired", "Expired"],
        ]}
      />
    );
  if (props.field.type === "boolean")
    return (
      <label>
        <input
          type="checkbox"
          checked={Boolean(props.value)}
          onChange={(e) => props.setValue(e.target.checked)}
        />{" "}
        {props.field.label}
      </label>
    );
  return (
    <label>
      {props.field.label}
      <input
        type={
          props.field.type === "date"
            ? "date"
            : props.field.type === "money"
            ? "number"
            : "text"
        }
        step={props.field.type === "money" ? "0.01" : undefined}
        value={
          props.field.type === "date"
            ? dateInput(props.value)
            : props.value || ""
        }
        readOnly={props.field.readOnly}
        onChange={(e) => props.setValue(e.target.value)}
      />
    </label>
  );
}
function readArray(key: string, fallback: string[]): string[] {
  try {
    const stored = window.sessionStorage.getItem(key);
    return stored ? JSON.parse(stored) : fallback;
  } catch {
    return fallback;
  }
}
function store(key: string, value: string[]): void {
  try {
    window.sessionStorage.setItem(key, JSON.stringify(value));
  } catch {}
}
function unique(values: string[]): string[] {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) =>
    a.localeCompare(b)
  );
}
function getColumn(key: string): Column {
  return col.filter((item) => item.key === key)[0];
}
function display(item: Intake, column: Column, base: string): string {
  const value = column.value(item);
  if (column.type === "link") return value ? "View Document" : "—";
  if (value === undefined || value === null || value === "") return "—";
  if (column.type === "money") return money.format(num(value));
  if (column.type === "date") return date(value);
  return String(value);
}
function nextSort(
  current: { key: string; desc: boolean },
  key: string
): { key: string; desc: boolean } {
  return { key, desc: current.key === key ? !current.desc : false };
}
function compare(
  left: Intake,
  right: Intake,
  sort: { key: string; desc: boolean },
  base: string
): number {
  const direction = sort.desc ? -1 : 1;
  if (sort.key !== "resident") {
    const column = getColumn(sort.key);
    if (column.type === "money") {
      return direction * (num(column.value(left)) - num(column.value(right)));
    }
    if (column.type === "date") {
      const leftDate = calendarDay(column.value(left));
      const rightDate = calendarDay(column.value(right));
      if (leftDate === undefined && rightDate === undefined) return 0;
      if (leftDate === undefined) return 1;
      if (rightDate === undefined) return -1;
      return direction * (leftDate - rightDate);
    }
  }
  const leftValue =
    sort.key === "resident"
      ? name(left)
      : display(left, getColumn(sort.key), base);
  const rightValue =
    sort.key === "resident"
      ? name(right)
      : display(right, getColumn(sort.key), base);
  return (
    direction *
    leftValue.localeCompare(rightValue, undefined, { numeric: true })
  );
}
function leaseMatches(item: Intake, status: string): boolean {
  const value = item.HAPContractEnd;
  if (!value) return false;
  const diff = new Date(value).getTime() - Date.now();
  return status === "expired" ? diff < 0 : diff >= 0 && diff <= 7776000000;
}
function dateWarningMessage(item: Intake, columnKey: string): string | undefined {
  const day = 24 * 60 * 60 * 1000;
  if (columnKey === "hapStart") {
    const hapStart = calendarDay(item.HAPContractStart);
    const moveIn = calendarDay(
      item.TenantName && item.TenantName.moveInDate
    );
    return hapStart !== undefined &&
      moveIn !== undefined &&
      Math.abs(hapStart - moveIn) > 30 * day
      ? "HAP Start is more than 30 days from the Move-in date."
      : undefined;
  }
  if (columnKey === "hapEnd") {
    const hapStart = calendarDay(item.HAPContractStart);
    const hapEnd = calendarDay(item.HAPContractEnd);
    if (hapStart === undefined || hapEnd === undefined) return undefined;
    const start = new Date(hapStart);
    const expectedEnd = Date.UTC(
      start.getUTCFullYear() + 1,
      start.getUTCMonth(),
      start.getUTCDate()
    );
    return Math.abs(hapEnd - expectedEnd) > 30 * day
      ? "HAP End is more than 30 days from the expected one-year end date."
      : undefined;
  }
  return undefined;
}
function calendarDay(value: any): number | undefined {
  if (!value) return undefined;
  const parsed = new Date(value);
  return isFinite(parsed.getTime())
    ? Date.UTC(
        parsed.getUTCFullYear(),
        parsed.getUTCMonth(),
        parsed.getUTCDate()
      )
    : undefined;
}
function calculateDashboardCharts(records: Intake[]): {
  snapshot: SnapshotMetric[];
  resident: ResidentMetricGroup[];
  properties: PropertyMetric[];
} {
  const now = Date.now();
  const inNinetyDays = now + 90 * 24 * 60 * 60 * 1000;
  const dateIsBetween = (value: any, start: number, end: number): boolean => {
    const timestamp = value ? new Date(value).getTime() : NaN;
    return isFinite(timestamp) && timestamp >= start && timestamp <= end;
  };
  const leaseEnd = (item: Intake): any =>
    item.TenantName && item.TenantName.leaseEndDate;
  const snapshot: SnapshotMetric[] = [
    {
      label: "HAP dates unknown",
      value: records.filter((item) => !item.HAPContractStart).length,
      color: "#1593a5",
    },
    {
      label: "DHA status unknown",
      value: records.filter(
        (item) => !String(item.DHAStatus || "").trim()
      ).length,
      color: "#6957d8",
    },
    {
      label: "Expired leases",
      value: records.filter((item) =>
        dateIsBetween(leaseEnd(item), 0, now - 1)
      ).length,
      color: "#d84b4b",
    },
    {
      label: "Leases expiring ≤90 days",
      value: records.filter((item) =>
        dateIsBetween(leaseEnd(item), now, inNinetyDays)
      ).length,
      color: "#e59a17",
    },
    {
      label: "DHA program expiring ≤90 days",
      value: records.filter((item) =>
        dateIsBetween(item.HAPContractEnd, now, inNinetyDays)
      ).length,
      color: "#18a66a",
    },
  ];
  const today = calendarDay(new Date()) || now;
  const leaseLimit = today + 90 * 24 * 60 * 60 * 1000;
  const resident: ResidentMetricGroup[] = [
    buildResidentMetric("DHA status", records, (item) =>
      String(item.DHAStatus || "").trim() || "Unknown"
    ),
    buildResidentMetric("Residency status", records, (item) =>
      String(residency(item) || "").trim() || "Unknown"
    ),
    buildResidentMetric("Lease health", records, (item) => {
      const end = calendarDay(
        item.TenantName && item.TenantName.leaseEndDate
      );
      if (end === undefined) return "Unknown";
      if (end < today) return "Expired";
      if (end <= leaseLimit) return "Expiring ≤90 days";
      return "Active >90 days";
    }),
    buildResidentMetric("Balance position", records, (item) => {
      const balance = num(item.Balance);
      if (balance > 0) return "Balance owing";
      if (balance < 0) return "Credit balance";
      return "Zero balance";
    }),
  ];
  const propertyLookup: { [name: string]: PropertyMetric } = {};
  records.forEach((item) => {
    const propertyName = prop(item).trim() || "Unknown property";
    if (!propertyLookup[propertyName]) {
      propertyLookup[propertyName] = {
        name: propertyName,
        residents: 0,
        rent: 0,
      };
    }
    propertyLookup[propertyName].residents += 1;
    propertyLookup[propertyName].rent += num(item.MonthlyRent);
  });
  const properties = Object.keys(propertyLookup)
    .map((propertyName) => propertyLookup[propertyName])
    .sort(
      (left, right) =>
        right.residents - left.residents ||
        left.name.localeCompare(right.name, undefined, { sensitivity: "base" })
    );
  return { snapshot, resident, properties };
}
function buildResidentMetric(
  label: string,
  records: Intake[],
  category: (item: Intake) => string
): ResidentMetricGroup {
  const counts: { [category: string]: number } = {};
  records.forEach((item) => {
    const value = category(item);
    counts[value] = (counts[value] || 0) + 1;
  });
  const segments = Object.keys(counts)
    .map((value) => ({
      label: value,
      value: counts[value],
      color: residentMetricColor(value),
    }))
    .sort(
      (left, right) =>
        right.value - left.value || left.label.localeCompare(right.label)
    );
  return { label, segments };
}
function residentMetricColor(label: string): string {
  const value = label.toLowerCase();
  if (value.indexOf("unknown") >= 0) return "#6957d8";
  if (value.indexOf("expired") >= 0 || value.indexOf("owing") >= 0)
    return "#d84b4b";
  if (value.indexOf("expiring") >= 0 || value.indexOf("pending") >= 0)
    return "#e59a17";
  if (
    value.indexOf("active") >= 0 ||
    value.indexOf("current") >= 0 ||
    value.indexOf("approved") >= 0 ||
    value.indexOf("zero") >= 0
  )
    return "#18a66a";
  if (value.indexOf("credit") >= 0) return "#1767a3";
  return "#1593a5";
}
function tenantLabel(person: Person): string {
  return (
    person.shortName ||
    person.personName ||
    person.Title ||
    `Tenant ${person.Id}`
  );
}
function tenantDisplay(id: number, people: Person[]): string {
  const found = people.filter((person) => person.Id === id)[0];
  return found ? tenantLabel(found) : "";
}
function ledgerUnit(item: LedgerEntry, people: Person[]): string {
  const tenant = people.filter((person) => person.Id === item.TenantNameId)[0];
  return tenant ? tenant.unitName || "" : "";
}
function ledgerProperty(item: LedgerEntry, people: Person[]): string {
  const tenant = people.filter((person) => person.Id === item.TenantNameId)[0];
  return tenant ? tenant.propertyName || "" : "";
}
function currentLedgerMonthKey(): string {
  const current = new Date();
  const month = current.getUTCMonth() + 1;
  return `${current.getUTCFullYear()}-${month < 10 ? `0${month}` : month}`;
}
function monthLabel(monthKey: string): string {
  const parts = monthKey.split("-");
  return new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, 1))
    .toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" });
}
function ledgerMonthOptions(): Array<{ key: string; label: string }> {
  const current = new Date();
  const options: Array<{ key: string; label: string }> = [];
  for (let offset = 0; offset >= -120; offset--) {
    const dateValue = new Date(
      Date.UTC(current.getUTCFullYear(), current.getUTCMonth() + offset, 1)
    );
    const month = dateValue.getUTCMonth() + 1;
    const key = `${dateValue.getUTCFullYear()}-${month < 10 ? `0${month}` : month}`;
    options.push({ key, label: monthLabel(key) });
  }
  return options;
}
function ledgerMonthRange(monthKey: string): { start: string; end: string } {
  const parts = monthKey.split("-");
  const year = Number(parts[0]);
  const monthIndex = Number(parts[1]) - 1;
  return {
    start: new Date(Date.UTC(year, monthIndex, 1)).toISOString(),
    end: new Date(Date.UTC(year, monthIndex + 1, 1)).toISOString(),
  };
}
function ledgerDetailColumn(
  key: LedgerDetailSortKey,
  label: string,
  sort: { key: LedgerDetailSortKey; desc: boolean },
  onSort: (key: LedgerDetailSortKey) => void
): React.ReactElement {
  const active = sort.key === key;
  return (
    <th
      className={
        key === "date"
          ? styles.ledgerDetailDate
          : key === "transactionType"
          ? styles.ledgerDetailType
          : undefined
      }
      aria-sort={active ? (sort.desc ? "descending" : "ascending") : "none"}
    >
      <button type="button" onClick={() => onSort(key)}>
        {label}
        <span aria-hidden="true">{active ? (sort.desc ? "↓" : "↑") : "↕"}</span>
      </button>
    </th>
  );
}
function compareLedgerTransactions(
  left: LedgerEntry,
  right: LedgerEntry,
  sort: { key: LedgerDetailSortKey; desc: boolean }
): number {
  const leftValue = left[sort.key];
  const rightValue = right[sort.key];
  let result: number;
  if (sort.key === "date" || sort.key === "DateReversed") {
    result =
      new Date(leftValue || 0).getTime() -
      new Date(rightValue || 0).getTime();
  } else if (sort.key === "Amount") {
    result = num(leftValue) - num(rightValue);
  } else {
    result = String(leftValue || "").localeCompare(
      String(rightValue || ""),
      undefined,
      { sensitivity: "base" }
    );
  }
  return sort.desc ? -result : result;
}
function ledgerValueTotal(values: { [category: string]: number }): number {
  return Object.keys(values).reduce(
    (total, category) => total + values[category],
    0
  );
}
function ledgerNetBalance(row: LedgerPivotRow): number {
  return (
    ledgerValueTotal(row.charges) -
    ledgerValueTotal(row.payments) -
    ledgerValueTotal(row.credits)
  );
}
function buildLedgerPivot(
  entries: LedgerEntry[],
  people: Person[]
): LedgerPivot {
  const chargeCategories: { [category: string]: boolean } = {};
  const paymentCategories: { [category: string]: boolean } = {};
  const creditCategories: { [category: string]: boolean } = {};
  const groups: { [key: string]: LedgerPivotRow } = {};
  entries.forEach((entry) => {
    const transactionType = (entry.transactionType || "").trim().toLowerCase();
    const category = (entry.CategoryName || "Uncategorized").trim();
    const entryDate = entry.date ? new Date(entry.date) : undefined;
    if (!entryDate || isNaN(entryDate.getTime())) return;
    const monthNumber = entryDate.getUTCMonth() + 1;
    const monthKey = `${entryDate.getUTCFullYear()}-${
      monthNumber < 10 ? `0${monthNumber}` : monthNumber
    }`;
    const month = entryDate.toLocaleDateString("en-US", {
      month: "long",
      year: "numeric",
      timeZone: "UTC",
    });
    const tenant =
      tenantDisplay(Number(entry.TenantNameId || 0), people) ||
      "Unknown tenant";
    const property = ledgerProperty(entry, people);
    const unit = ledgerUnit(entry, people);
    const key = JSON.stringify([tenant, property, unit, monthKey]);
    if (!groups[key]) {
      groups[key] = {
        key,
        tenant,
        property,
        unit,
        monthKey,
        month,
        charges: {},
        payments: {},
        credits: {},
        transactions: [],
      };
    }
    groups[key].transactions.push(entry);
    if (
      transactionType !== "charge" &&
      transactionType !== "payment" &&
      transactionType !== "credit"
    ) return;
    const values = transactionType === "charge"
      ? groups[key].charges
      : transactionType === "payment"
      ? groups[key].payments
      : groups[key].credits;
    values[category] = (values[category] || 0) + num(entry.Amount);
    if (transactionType === "charge") chargeCategories[category] = true;
    else if (transactionType === "payment") paymentCategories[category] = true;
    else creditCategories[category] = true;
  });
  return {
    chargeCategories: Object.keys(chargeCategories).sort(),
    paymentCategories: Object.keys(paymentCategories).sort(),
    creditCategories: Object.keys(creditCategories).sort(),
    rows: Object.keys(groups)
      .map((key) => groups[key])
      .filter(
        (row) =>
          Object.keys(row.charges).length > 0 ||
          Object.keys(row.payments).length > 0 ||
          Object.keys(row.credits).length > 0
      )
      .sort(
        (left, right) =>
          right.monthKey.localeCompare(left.monthKey) ||
          left.tenant.localeCompare(right.tenant) ||
          left.property.localeCompare(right.property) ||
          left.unit.localeCompare(right.unit)
      ),
  };
}
function calculateKpis(
  records: Intake[],
  base: string
): Array<{ title: string; value: string; subtitle: string }> {
  const rent = records.reduce(
    (total, item) => total + num(item.MonthlyRent),
    0
  );
  const hap = records.reduce((total, item) => total + num(item.HAPPortion), 0);
  const tenant = records.reduce(
    (total, item) => total + num(item.HAPAmendment1Portion),
    0
  );
  const owing = records.filter((item) => num(item.Balance) > 0);
  const balance = owing.reduce((total, item) => total + num(item.Balance), 0);
  const highest = owing.sort((a, b) => num(b.Balance) - num(a.Balance))[0];
  const docs = records.filter((item) => Boolean(item.Id)).length;
  const invoices = records.filter((item) => Boolean(item.WelcomeLetter)).length;
  const hapDates = records.filter((item) =>
    Boolean(item.HAPContractStart && item.HAPContractEnd)
  ).length;
  const completeness = records.length
    ? Math.round(((docs + invoices + hapDates) / (records.length * 3)) * 100)
    : 0;
  return [
    {
      title: "Residents Shown",
      value: String(records.length),
      subtitle: "Global filters applied",
    },
    {
      title: "Monthly Rent Exposure",
      value: money.format(rent),
      subtitle: "Globally filtered residents",
    },
    {
      title: "HAP Subsidy",
      value: money.format(hap),
      subtitle: `${
        rent ? ((hap / rent) * 100).toFixed(1) : "0.0"
      }% of rent covered`,
    },
    {
      title: "Tenant Portion",
      value: money.format(tenant),
      subtitle: "Calculated responsibility",
    },
    {
      title: "Accounts Owing Balance",
      value: `${owing.length} Residents`,
      subtitle: "Residents with balance > $0",
    },
    {
      title: "Balance Owed",
      value: money.format(balance),
      subtitle: "Total positive resident balances",
    },
    {
      title: "Highest Balance Owed",
      value: money.format(num(highest && highest.Balance)),
      subtitle: highest ? name(highest) : "No outstanding balances",
    },
    {
      title: "Document Completeness",
      value: `${completeness}%`,
      subtitle: "Program docs, invoices, HAP dates",
    },
  ];
}
