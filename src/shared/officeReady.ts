export async function officeReady(): Promise<void> {
  if (typeof Office === "undefined") return;
  await Office.onReady();
}

