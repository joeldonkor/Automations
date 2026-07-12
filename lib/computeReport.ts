import type { GroupTotals, PublisherRow } from "./types";

function sum(values: number[]): number {
  return values.reduce((a, b) => a + b, 0);
}

export function computeGroupTotals(rows: PublisherRow[]): GroupTotals {
  const active = rows.filter((r) => !r.irregular);
  const publishers = active.filter((r) => r.category === "publisher");
  const auxPioneers = active.filter((r) => r.category === "auxPioneer");
  const regularPioneers = active.filter((r) => r.category === "regularPioneer");

  return {
    // Active publishers = Aux pioneers + Regular pioneers + Irregular publishers + Publishers,
    // i.e. everyone on the roster — irregular ones are still active, just not reporting consistently.
    activePublishers: rows.length,
    publishersReported: publishers.filter((r) => r.reported).length,
    publisherBibleStudies: sum(publishers.map((r) => r.bibleStudies)),
    auxPioneers: auxPioneers.length,
    auxPioneerHours: sum(auxPioneers.map((r) => r.hours ?? 0)),
    auxBibleStudies: sum(auxPioneers.map((r) => r.bibleStudies)),
    regularPioneers: regularPioneers.length,
    regularPioneerHours: sum(regularPioneers.map((r) => r.hours ?? 0)),
    regularBibleStudies: sum(regularPioneers.map((r) => r.bibleStudies)),
    irregularPublishers: rows.length - active.length,
  };
}
