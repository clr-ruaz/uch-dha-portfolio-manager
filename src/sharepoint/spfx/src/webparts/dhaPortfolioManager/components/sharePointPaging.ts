export async function readAllSharePointItems(
  url: string,
  get: (url: string) => Promise<any>
): Promise<any[]> {
  const items: any[] = [];
  let nextUrl = url;
  while (nextUrl) {
    const result = await get(nextUrl);
    items.push(...(result.value || (result.d && result.d.results) || []));
    nextUrl = result["@odata.nextLink"] || result["odata.nextLink"] ||
      (result.d && result.d.__next) || "";
  }
  return items;
}
