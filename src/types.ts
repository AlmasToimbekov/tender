interface BasicInfo {
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
    };
  };
  primaryOKED: {
    value: string;
  };
  secondaryOKED: {
    value: string[];
  };
  addressRu: {
    value: string;
  };
}

interface GosZakupContacts {
  phone: {
    value: string;
    href: string;
  }[];
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
