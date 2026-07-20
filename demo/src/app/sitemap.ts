import type { MetadataRoute } from "next";
import { SITE_URL } from "@/lib/constants";

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: SITE_URL,
      lastModified: "2026-07-20",
      changeFrequency: "monthly",
      priority: 1,
    },
    {
      url: `${SITE_URL}/docs`,
      lastModified: "2026-07-20",
      changeFrequency: "monthly",
      priority: 0.8,
    },
  ];
}
