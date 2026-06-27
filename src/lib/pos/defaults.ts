import type { Service } from "@/types/pos";

export const DEFAULT_SERVICES: Service[] = [
  { id: "xerox-single", name: "Xerox", category: "Xerox", optionLabel: "Single side", price: 3, sortOrder: 10 },
  { id: "xerox-back", name: "Xerox", category: "Xerox", optionLabel: "Back to back", price: 6, sortOrder: 20 },
  { id: "print-bw-short", name: "Black and white print", category: "Printing", optionLabel: "Short", price: 5, sortOrder: 30 },
  { id: "print-bw-long", name: "Black and white print", category: "Printing", optionLabel: "Long", price: 7, sortOrder: 40 },
  { id: "print-bw-a4", name: "Black and white print", category: "Printing", optionLabel: "A4", price: 6, sortOrder: 50 },
  { id: "print-color-short", name: "Colored print", category: "Printing", optionLabel: "Short", price: 10, sortOrder: 60 },
  { id: "print-color-long", name: "Colored print", category: "Printing", optionLabel: "Long", price: 15, sortOrder: 70 },
  { id: "print-color-a4", name: "Colored print", category: "Printing", optionLabel: "A4", price: 15, sortOrder: 80 },
  { id: "gov-custom", name: "Government service", category: "Online Services", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 90 },
  { id: "nbi-custom", name: "NBI assistance", category: "Online Services", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 100 },
  { id: "police-clearance-custom", name: "Police clearance", category: "Online Services", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 110 },
  { id: "psa-custom", name: "PSA assistance", category: "Online Services", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 120 },
  { id: "laminating-custom", name: "Laminating", category: "Finishing", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 130 },
  { id: "misc-custom", name: "Other shop service", category: "Custom", optionLabel: "Custom price", price: 0, isCustomPrice: true, sortOrder: 140 },
];
