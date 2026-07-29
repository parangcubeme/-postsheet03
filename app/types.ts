export type Row = Record<string, unknown>;
export type MarketTab = "smartstore" | "esm";
export type StepId = "category" | "notice" | "origin" | "shipping" | "price" | "review";

export type Product = {
  id: string;
  raw: Row;
  productName: string;
  sellerCode: string;
  basePrice: number;
  finalPrice: number;
  stock: number;
  mainImage: string;
  additionalImage: string;
  detailHtml: string;
  shippingFee: number;
  vatType: string;
  originDirect: string;
  categoryGroup: string;
  categoryTemplateCode: string;
  categoryCode: string;
  auctionExposureCode: string;
  gmarketExposureCode: string;
  productGroupCode: string;
  noticeTemplateCode: string;
  originProductType: string;
  originRegionType: string;
  originRegionCode: string;
  multipleOrigins: string;
  shippingGroup: string;
  departureCode: string;
  shippingPolicyNumber: string;
  returnAddressCode: string;
  auctionShippingPolicy: string;
  gmarketShippingPolicy: string;
  courierCode: string;
  returnShippingFee: number;
};

export type ProductGroup = {
  name: string;
  items: Product[];
};

export type DownloadPage = {
  key: string;
  category: string;
  page: number;
  items: Product[];
};

export type WorkerSuccess = {
  ok: true;
  rows: Row[];
  products: Product[];
  sheetName: string;
  headerRow: number;
};

export type WorkerFailure = {
  ok: false;
  message: string;
};

export type WorkerResponse = WorkerSuccess | WorkerFailure;
