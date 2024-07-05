export interface BasicInfo {
  isDeleted: boolean;
  registrationDate: {
    value: string | null;
  };
  onMarket: {
    years: number;
    months: number;
  } | null;
  ceo: {
    value: {
      title: string;
    } | null;
  };
  primaryOKED: {
    value: string;
  };
  secondaryOKED: {
    value: string[] | null;
  };
  addressRu: {
    value: string;
  };
}

interface GosZakupContacts {
  phone:
    | {
        value: string;
        href: string;
      }[]
    | null;
  website: string | null;
  email: {
    value: string;
    href: string;
  }[];
}

export interface CompanyFullInfo {
  basicInfo: BasicInfo;
  gosZakupContacts: GosZakupContacts | null;
}
