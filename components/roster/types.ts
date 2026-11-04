export type Gender = 'male' | 'female';
export type CadetStatus = 'active' | 'graduated' | 'expelled';

export interface StagingCadetDTO {
  temp_id: string;
  id?: number;
  rank: string;
  fio: string;
  short_fio: string;
  group_num: string;
  course: number;
  gender: Gender;
  status: CadetStatus;
  needs_manual_review: boolean;
  warnings: string[];
  sort_order: number;
}

export interface ParsePreviewResponse {
  cadets: StagingCadetDTO[];
  total_count: number;
  graduated_count: number;
  needs_review_count: number;
  males_count: number;
  females_count: number;
}

export interface RosterSliceBalance {
  slice_name: 'total' | 'male' | 'female';
  listed_count: number;
  present_count: number;
  absent_count: number;
  categories_breakdown: Record<string, number>;
  is_balanced: boolean;
  discrepancy: number;
  error_message?: string | null;
}

export interface TriSliceBalance {
  total: RosterSliceBalance;
  male: RosterSliceBalance;
  female: RosterSliceBalance;
  is_sum_balanced: boolean;
  error_message?: string | null;
}

export interface Cadet {
  id: number;
  rank: string;
  fio: string;
  short_fio: string;
  group_num: string;
  course: number;
  gender: Gender | 'М' | 'Ж';
  status: CadetStatus;
  sort_order: number;
}

export interface AbsenceRecord {
  id: number;
  cadet_id: number;
  rank: string;
  fio: string;
  short_fio: string;
  group_num: string;
  course: number;
  gender: Gender | 'М' | 'Ж';
  type_id: number;
  type_name: string;
  type_code?: string;
  category: string;
  note?: string;
}
