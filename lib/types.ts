export type PublisherCategory = "publisher" | "auxPioneer" | "regularPioneer";

export type PublisherRow = {
  name: string;
  reported: boolean;
  bibleStudies: number;
  hours: number | null;
  notes: string;
  irregular: boolean;
  category: PublisherCategory;
};

export type GroupData = {
  label: string;
  rows: PublisherRow[];
};

export type GroupTotals = {
  activePublishers: number;
  publishersReported: number;
  publisherBibleStudies: number;
  auxPioneers: number;
  auxPioneerHours: number;
  auxBibleStudies: number;
  regularPioneers: number;
  regularPioneerHours: number;
  regularBibleStudies: number;
  irregularPublishers: number;
};

export const REPORT_ROWS: { label: string; key: keyof GroupTotals }[] = [
  { label: "Active Publishers", key: "activePublishers" },
  { label: "Publishers Number of Report", key: "publishersReported" },
  { label: "Publishers Bible Studies", key: "publisherBibleStudies" },
  { label: "Auxiliary Pioneers", key: "auxPioneers" },
  { label: "Auxiliary Pioneers Hours", key: "auxPioneerHours" },
  { label: "Auxiliary Bible Studies", key: "auxBibleStudies" },
  { label: "Regular Pioneers", key: "regularPioneers" },
  { label: "Regular Pioneers Hours", key: "regularPioneerHours" },
  { label: "Regular Bible Studies", key: "regularBibleStudies" },
  { label: "Irregular Publishers", key: "irregularPublishers" },
];
