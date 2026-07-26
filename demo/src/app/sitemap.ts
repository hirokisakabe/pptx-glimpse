import type { MetadataRoute } from "next";
import type { PageMapItem } from "nextra";
import { getPageMap } from "nextra/page-map";
import { SITE_URL } from "@/lib/constants";
import { DOCS_NAVIGATION } from "@/lib/docs-navigation";

export default async function sitemap(): Promise<MetadataRoute.Sitemap> {
  const documentationRoutes = DOCS_NAVIGATION.flatMap((group) =>
    group.items.map((item) => item.href),
  );
  const apiRoutes = collectRoutes(await getPageMap("/docs/api")).filter((route) =>
    route.startsWith("/docs/api"),
  );
  const routes = [...new Set([...documentationRoutes, ...apiRoutes])];

  return [
    {
      url: SITE_URL,
      lastModified: "2026-07-25",
      changeFrequency: "monthly",
      priority: 1,
    },
    ...routes.map((route) => ({
      url: `${SITE_URL}${route}`,
      lastModified: "2026-07-25",
      changeFrequency: "monthly" as const,
      priority: route === "/docs" ? 0.8 : 0.7,
    })),
  ];
}

function collectRoutes(items: readonly PageMapItem[]): string[] {
  return items.flatMap((item) => {
    if (!("route" in item)) return [];
    const children = "children" in item ? collectRoutes(item.children) : [];
    return [item.route, ...children];
  });
}
