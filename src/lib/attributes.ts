// Yivi/SIDN attribute schemes supported by the PostGuard policy editor.

export const EMAIL_ATTRIBUTE_TYPE = "pbdf.sidn-pbdf.email.email";

export interface AttributeDescriptor {
  type: string;
  defaultLabel: string;
}

export const SUPPORTED_ATTRIBUTES: AttributeDescriptor[] = [
  { type: EMAIL_ATTRIBUTE_TYPE, defaultLabel: "Email address" },
  { type: "pbdf.sidn-pbdf.mobilenumber.mobilenumber", defaultLabel: "Mobile number" },
  { type: "pbdf.gemeente.personalData.surname", defaultLabel: "Surname" },
  { type: "pbdf.gemeente.personalData.dateofbirth", defaultLabel: "Date of birth" },
];
