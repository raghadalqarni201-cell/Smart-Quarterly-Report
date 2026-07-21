export interface ExcelRecord {
  category: string; // Status or Rejection Reason
  cases: number;
  amount: number;
}

export interface ClinicData {
  statusTable: ExcelRecord[];
  rejectionTable: ExcelRecord[];
}

export interface MonthlyReportPayload {
  monthIndex: number; // 0, 1, or 2 (representing Month 1, 2, 3)
  dateRange: string;  // e.g. "2026-01-01 to 2026-01-31" or "January 2026"
  insuranceCompany: string;
  clinics: {
    [rawName: string]: ClinicData;
  };
}

export interface ClinicMetadata {
  rawName: string;
  shortName: string;
  iconName: string; // lucide icon name
  description: string;
}

export const CLINICS_METADATA: ClinicMetadata[] = [
  {
    rawName: "Al Osrah Medical Center in Hijaz of Royal commission in Jubail Industrial",
    shortName: "Al-Hejaz",
    iconName: "ShieldAlert",
    description: "Al Osrah Medical Center (Hijaz RC Jubail)"
  },
  {
    rawName: "Al Dafi first aid Medical Center of Royal commission in Jubail Industrial",
    shortName: "Al-Dafi",
    iconName: "Flame",
    description: "Al Dafi First Aid Medical Center"
  },
  {
    rawName: "Al Howilat first aid Medical Center of Royal commission in Jubail Industrial",
    shortName: "Al-Huwaylat",
    iconName: "HeartPulse",
    description: "Al Howilat First Aid Medical Center"
  },
  {
    rawName: "Al Farouq first aid Medical Center of Royal commission in Jubail Industrial",
    shortName: "Al-Farouq",
    iconName: "Activity",
    description: "Al Farouq First Aid Medical Center"
  },
  {
    rawName: "Jalmoud first aid Medical Center of Royal commission in Jubail Industrial",
    shortName: "Jalmudah",
    iconName: "Stethoscope",
    description: "Jalmoud First Aid Medical Center"
  },
  {
    rawName: "Royal Commission Hospital in Jubail",
    shortName: "RCH-Jubail",
    iconName: "Hospital",
    description: "Royal Commission Hospital (Jubail Main)"
  },
  {
    rawName: "Ras Al Khair first aid Medical Center of Royal commission in Jubail Industrial",
    shortName: "Ras-Al-Khair",
    iconName: "Anchor",
    description: "Ras Al Khair First Aid Medical Center"
  }
];

export const INSURANCE_COMPANIES = [
  "Bupa Arabia",
  "Tawuniya",
  "Medgulf",
  "Malath Insurance",
  "Al Rajhi Takaful",
  "AXA Cooperative",
  "GlobeMed Saudi",
  "Allianz SF",
  "UCA",
  "GIG Gulf",
  "Other Insurance / Mixed"
];
