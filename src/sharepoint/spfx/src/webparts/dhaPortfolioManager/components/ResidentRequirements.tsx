import * as React from "react";
import styles from "./Dashboard.module.scss";

const requirementGroups = [
  {
    title: "Lease Details",
    fields: [
      { field: "MonthlyRent", label: "Monthly Rent" },
      { field: "HAPPortion", label: "HAP Portion" },
      { field: "ApplicationFee", label: "Application Fee" },
      { field: "AdministrativeFee", label: "Admin Fee" },
      { field: "RiskFee", label: "Risk Fee" },
      { field: "SecurityDeposit", label: "Security Deposit" },
      { field: "LandlordIncentiveFee", label: "Landlord Incentive" },
    ],
  },
  {
    title: "DHA Program",
    fields: [
      { field: "CaseworkerName", label: "Caseworker Name" },
      { field: "CaseworkerContact", label: "Caseworker Contact" },
      { field: "HAPContractStart", label: "Contract Start" },
      { field: "HAPContractEnd", label: "Contract End" },
    ],
  },
  {
    title: "Documents",
    fields: [
      { field: "HAPContract", label: "DHA Contract" },
      { field: "WelcomeLetter", label: "Welcome Letter" },
    ],
  },
];

export const requirementFields = requirementGroups.reduce<string[]>(
  (fields, group) => fields.concat(group.fields.map((item) => item.field)),
  []
);

// SharePoint Yes/No fields are booleans. Only an explicit No opts out;
// older records with missing values retain their existing requirements.
export const isRequired = (value: unknown): boolean => value !== false;

export function isFieldNotApplicable(
  field: string,
  draft: { [key: string]: unknown }
): boolean {
  if (/^HAPAmendment[1-3](Date|Portion)$/.test(field)) {
    // Amendment applicability is display-only and follows both contract dates.
    return !isRequired(draft.HAPContractStartRequired) ||
      !isRequired(draft.HAPContractEndRequired);
  }
  return requirementFields.indexOf(field) >= 0 &&
    !isRequired(draft[`${field}Required`]);
}

export function requirementUpdates(
  draft: { [key: string]: unknown },
  original: { [key: string]: unknown }
): { [key: string]: boolean } {
  const updates: { [key: string]: boolean } = {};
  requirementFields.forEach((field) => {
    const key = `${field}Required`;
    if (isRequired(draft[key]) !== isRequired(original[key])) {
      updates[key] = isRequired(draft[key]);
    }
  });
  return updates;
}

export default function ResidentRequirements(props: {
  draft: { [key: string]: unknown };
  saving: boolean;
  onChange: (key: string, required: boolean) => void;
}): React.ReactElement {
  const count = requirementFields.filter((field) =>
    isRequired(props.draft[`${field}Required`])
  ).length;

  return (
    <details className={styles.residentRequirements}>
      <summary className={styles.requirementsSummary}>
        <svg
          className={styles.requirementsChevron}
          width="16"
          height="16"
          viewBox="0 0 16 16"
          aria-hidden="true"
          focusable="false"
        >
          <path d="m6 3 5 5-5 5" fill="none" stroke="currentColor" strokeWidth="1.75" />
        </svg>
        <span className={styles.requirementsTitle}>
          Required information &amp; documents
        </span>
        <span className={styles.requirementsCount}>
          {count} of {requirementFields.length} required
        </span>
      </summary>
      <div className={styles.requirementsBody}>
        <p className={styles.requirementsDescription}>
          Select the information and documents required for this resident.
        </p>
        <div className={styles.requirementsGroups}>
          {requirementGroups.map((group) => (
            <fieldset key={group.title} disabled={props.saving}>
              <legend>{group.title}</legend>
              {group.fields.map((item) => {
                const key = `${item.field}Required`;
                return (
                  <label key={key} className={styles.requirementOption}>
                    <input
                      type="checkbox"
                      checked={isRequired(props.draft[key])}
                      onChange={(event) => props.onChange(key, event.target.checked)}
                    />
                    <span>{item.label}</span>
                  </label>
                );
              })}
            </fieldset>
          ))}
        </div>
      </div>
    </details>
  );
}
