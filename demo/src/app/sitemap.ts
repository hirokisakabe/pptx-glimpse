import type { MetadataRoute } from "next";
import { SITE_URL } from "@/lib/constants";

export default function sitemap(): MetadataRoute.Sitemap {
  return [
    {
      url: SITE_URL,
      lastModified: "2025-03-05",
      changeFrequency: "monthly",
      priority: 1,
    },
  ];
}
