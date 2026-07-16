import { Suspense } from "react";
import { Canvas, useLoader } from "@react-three/fiber";
import { OrbitControls, Center, Html, useProgress, Bounds } from "@react-three/drei";
import { STLLoader } from "three/examples/jsm/loaders/STLLoader.js";

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

function StlMesh({ url }: { url: string }) {
  const geometry = useLoader(STLLoader, url);
  return (
    <Center>
      <mesh geometry={geometry} castShadow receiveShadow rotation={[-Math.PI / 2, 0, 0]}>
        <meshStandardMaterial color="#9aa5ff" metalness={0.25} roughness={0.45} />
      </mesh>
    </Center>
  );
}

export default function StlViewer({ url }: { url: string }) {
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
            <StlMesh url={url} />
          </Bounds>
        </Suspense>
        <gridHelper args={[20, 20, "#26304a", "#131a2c"]} position={[0, -1.001, 0]} />
        <OrbitControls makeDefault enableDamping dampingFactor={0.08} />
      </Canvas>
    </div>
  );
}
