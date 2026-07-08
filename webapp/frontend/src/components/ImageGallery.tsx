import { useState } from "react";
import { X, ZoomIn } from "lucide-react";
import type { AssetRef } from "../api/client";

export default function ImageGallery({ images }: { images: AssetRef[] }) {
  const [active, setActive] = useState<AssetRef | null>(null);

  return (
    <>
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
        {images.map((img) => (
          <button
            key={img.name}
            onClick={() => setActive(img)}
            className="group relative aspect-square overflow-hidden rounded-xl bg-ink-800 ring-1 ring-ink-700/60 transition hover:ring-brand-500/60"
          >
            <img src={img.url} alt={img.name} className="h-full w-full object-cover transition duration-300 group-hover:scale-105" />
            <div className="absolute inset-0 flex items-center justify-center bg-black/0 opacity-0 transition group-hover:bg-black/30 group-hover:opacity-100">
              <ZoomIn className="h-5 w-5 text-white" />
            </div>
            <div className="absolute bottom-0 left-0 right-0 truncate bg-gradient-to-t from-black/70 to-transparent px-2 py-1.5 text-left text-[10px] font-medium text-white">
              {prettyLabel(img.name)}
            </div>
          </button>
        ))}
      </div>

      {active && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-6 backdrop-blur-sm"
          onClick={() => setActive(null)}
        >
          <button
            className="absolute right-6 top-6 rounded-full bg-ink-800/80 p-2 text-ink-200 hover:bg-ink-700"
            onClick={() => setActive(null)}
          >
            <X className="h-5 w-5" />
          </button>
          <img
            src={active.url}
            alt={active.name}
            className="max-h-[85vh] max-w-[90vw] rounded-2xl object-contain shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          />
          <div className="absolute bottom-8 left-1/2 -translate-x-1/2 rounded-full bg-ink-900/85 px-4 py-1.5 text-xs font-medium text-ink-200">
            {prettyLabel(active.name)}
          </div>
        </div>
      )}
    </>
  );
}

function prettyLabel(name: string) {
  const stem = name.replace(/\.[^.]+$/, "");
  const parts = stem.split(/_FULL_?/i);
  return (parts[1] || stem).replace(/_/g, " ").trim() || stem;
}
