import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useState } from "react";
// @ts-expect-error - JSX component without types
import Minutario from "@/components/Minutario.jsx";

export const Route = createFileRoute("/")({
  component: Index,
  head: () => ({
    meta: [
      { title: "Minutário - Petição Inicial" },
      { name: "description", content: "Gerador de minutas de petição inicial" },
      { name: "theme-color", content: "#0f1b3d" },
    ],
    links: [
      { rel: "manifest", href: "/manifest.json" },
      { rel: "apple-touch-icon", href: "/icon-192.png" },
    ],
  }),
});

function Index() {
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);
  if (!mounted) return null;
  return <Minutario />;
}
