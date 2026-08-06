import { Suspense, useEffect, useMemo } from "react";
import { Canvas, useLoader } from "@react-three/fiber";
import { OrbitControls, Bounds, Html, useProgress } from "@react-three/drei";
import { STLLoader } from "three/examples/jsm/loaders/STLLoader.js";
import { DoubleSide, EdgesGeometry } from "three";
import { computeGeomStats, normalizeGeometryToInches, type GeomStats } from "../lib/stlGeometry";

// Same palette the shop's overlay viewer uses: A = blue, B = orange.
export const OVERLAY_COLORS = {
  a: { solid: "#3b82f6", edge: "#7dd3fc", opacity: 0.45 },
  b: { solid: "#ff8c42", edge: "#ffd199", opacity: 0.34 },
};

function LoadingOverlay() {
  const { progress } = useProgress();
  return (
    <Html center>
      <div className="whitespace-nowrap rounded-full bg-ink-900/90 px-4 py-2 text-xs font-medium text-ink-200 shadow-lg ring-1 ring-ink-600/60">
        Loading model... {Math.round(progress)}%
      </div>
    </Html>
  );
}

function OverlayMesh({
  url,
  side,
  onStats,
}: {
  url: string;
  side: "a" | "b";
  onStats?: (stats: GeomStats) => void;
}) {
  const geometry = useLoader(STLLoader, url);
  const col = OVERLAY_COLORS[side];

  const stats = useMemo(() => {
    // Both overlaid meshes must be in the same unit or one dwarfs the other.
    normalizeGeometryToInches(geometry);
    geometry.computeVertexNormals();
    return computeGeomStats(geometry);
  }, [geometry]);

  const edges = useMemo(() => new EdgesGeometry(geometry, 20), [geometry]);

  useEffect(() => {
    onStats?.(stats);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [stats]);

  return (
    <group position={[-stats.bcx, -stats.bcy, -stats.bcz]} rotation={[-Math.PI / 2, 0, 0]}>
      <mesh geometry={geometry}>
        <meshPhongMaterial
          color={col.solid}
          transparent
          opacity={col.opacity}
          side={DoubleSide}
          depthWrite={false}
          polygonOffset
          polygonOffsetFactor={1}
          polygonOffsetUnits={1}
        />
      </mesh>
      <lineSegments geometry={edges}>
        <lineBasicMaterial color={col.edge} transparent opacity={0.9} />
      </lineSegments>
    </group>
  );
}

export default function StlOverlayViewer({
  urlA,
  urlB,
  onStatsA,
  onStatsB,
}: {
  urlA?: string;
  urlB?: string;
  onStatsA?: (s: GeomStats) => void;
  onStatsB?: (s: GeomStats) => void;
}) {
  return (
    <div className="relative h-full w-full overflow-hidden rounded-2xl bg-gradient-to-b from-ink-850 to-ink-950">
      <Canvas camera={{ position: [4, 4, 6], fov: 45 }}>
        <color attach="background" args={["#070d18"]} />
        <ambientLight intensity={0.6} />
        <directionalLight position={[1, 2, 1]} intensity={0.9} />
        <directionalLight position={[-1, -1, -2]} intensity={0.45} />
        <Suspense fallback={<LoadingOverlay />}>
          <Bounds fit clip observe margin={1.3}>
            {urlA && <OverlayMesh url={urlA} side="a" onStats={onStatsA} />}
            {urlB && <OverlayMesh url={urlB} side="b" onStats={onStatsB} />}
          </Bounds>
        </Suspense>
        <OrbitControls makeDefault enableDamping dampingFactor={0.08} />
      </Canvas>
    </div>
  );
}
