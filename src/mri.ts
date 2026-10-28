/**
 * Extract the MRI suffix from a sender URL or return the input unchanged.
 * API messages use full URLs (e.g. ".../contacts/8:orgid:uuid"),
 * while member IDs are just the MRI suffix ("8:orgid:uuid").
 */
export function extractMriSuffix(senderMri: string): string {
  const contactsPrefix = "/contacts/";
  const contactsIndex = senderMri.lastIndexOf(contactsPrefix);
  if (contactsIndex >= 0) {
    return senderMri.slice(contactsIndex + contactsPrefix.length);
  }
  return senderMri;
}
