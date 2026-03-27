export interface PatientData {
  [key: string]: string;
}

export interface DashboardStats {
  totalPatients: number;
  patientsByClinic: { name: string; value: number }[];
  patientsByDate: { date: string; count: number }[];
  patientsByInsurance: { name: string; value: number }[];
}
