import { Suspense, useEffect } from "react";
import { Canvas, useLoader } from "@react-three/fiber";
import { OrbitControls, Center, Html, useProgress, Bounds } from "@react-three/drei";
import { STLLoader } from "three/examples/jsm/loaders/STLLoader.js";
import { normalizeGeometryToInches } from "../lib/stlGeometry";

export interface StlInfo {
  x: number;
  y: number;
  z: number;
  tris: number;
}

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

function StlMesh({ url, onInfo }: { url: string; onInfo?: (info: StlInfo) => void }) {
  const geometry = useLoader(STLLoader, url);

  // SolidWorks writes STL in millimetres on this machine and STL carries no
  // unit tag, so scale to inches before anything measures or frames it.
  normalizeGeometryToInches(geometry);

  useEffect(() => {
    if (!onInfo) return;
    geometry.computeBoundingBox();
    const box = geometry.boundingBox;
    const tris = geometry.index ? geometry.index.count / 3 : geometry.attributes.position.count / 3;
    if (box) {
      onInfo({
        x: box.max.x - box.min.x,
        y: box.max.y - box.min.y,
        z: box.max.z - box.min.z,
        tris: Math.round(tris),
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [geometry, url]);

  return (
    <Center>
      <mesh geometry={geometry} castShadow receiveShadow rotation={[-Math.PI / 2, 0, 0]}>
        <meshStandardMaterial color="#9aa5ff" metalness={0.25} roughness={0.45} />
      </mesh>
    </Center>
  );
}

export default function StlViewer({ url, onInfo }: { url: string; onInfo?: (info: StlInfo) => void }) {
  return (
    <div className="relative h-full w-full overflow-hidden rounded-2xl bg-gradient-to-b from-ink-850 to-ink-950">
      <Canvas shadows camera={{ position: [4, 4, 6], fov: 45 }}>
        <color attach="background" args={["#0a0e1a"]} />
        <hemisphereLight intensity={0.55} color="#c9d3ff" groundColor="#05070d" />
        <directionalLight
          position={[5, 8, 5]}
          intensity={1.4}
          castShadow
          shadow-mapSize={[1024, 1024]}
        />
        <directionalLight position={[-6, 3, -4]} intensity={0.4} color="#6d5bff" />
        <Suspense fallback={<LoadingOverlay />}>
          <Bounds fit clip observe margin={1.3}>
            <StlMesh url={url} onInfo={onInfo} />
          </Bounds>
        </Suspense>
        <gridHelper args={[20, 20, "#26304a", "#131a2c"]} position={[0, -1.001, 0]} />
        <OrbitControls makeDefault enableDamping dampingFactor={0.08} />
      </Canvas>
    </div>
  );
}
