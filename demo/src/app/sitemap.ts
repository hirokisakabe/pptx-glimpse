import type { MetadataRoute } from "next";
import { SITE_URL } from "@/lib/constants";
import { DOCS_NAVIGATION } from "@/lib/docs-navigation";

export default function sitemap(): MetadataRoute.Sitemap {
  const documentationRoutes = DOCS_NAVIGATION.flatMap((group) =>
    group.items.map((item) => item.href),
  );

  return [
    {
      url: SITE_URL,
      lastModified: "2026-07-25",
      changeFrequency: "monthly",
      priority: 1,
    },
    ...documentationRoutes.map((route) => ({
      url: `${SITE_URL}${route}`,
      lastModified: "2026-07-25",
      changeFrequency: "monthly" as const,
      priority: route === "/docs" ? 0.8 : 0.7,
    })),
  ];
}
