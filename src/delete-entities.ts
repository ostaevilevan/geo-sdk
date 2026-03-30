import "dotenv/config";
import {
  Graph,
  personalSpace,
  getSmartAccountWalletClient,
} from "@geoprotocol/geo-sdk";

// ── CONFIGURATION ─────────────────────────────────────────────────────────────
const TARGET_SPACE = "0c747ad3af58ed6b27221a256498068e";
const ROOT_SPACE   = "a19c345ab9866679b001d7d2138d88a1";
const GEO_API      = "https://testnet-api.geobrowser.io/graphql";

// ── DELETE MODE ───────────────────────────────────────────────────────────────
// Case 1 — "wrong-space"   : delete everything we own in the space
//                            (uploaded to wrong space, or full reset)
// Case 2 — "wrong-type"    : delete all instances of specific wrong typeIds
//                            + orphaned supporting entities
//                            + locally-created Type and Property entities
// Case 4 — "wrong-batch"   : delete a specific batch range by entity names
//                            + orphaned supporting entities
// Case 5 — "wrong-logic"   : same as wrong-batch — delete affected entities
//                            so they can be re-uploaded with corrected logic
//
// (Case 3 — wrong linked entity — is a relation fix, not a delete. Handle manually.)

const MODE = "wrong-space" as "wrong-space" | "wrong-type" | "wrong-batch" | "wrong-logic";

// ── MODE SETTINGS (configure per run) ─────────────────────────────────────────

// Case 2: wrong-type — typeIds to target
const WRONG_TYPE_IDS = new Set([
  "1362f6523665771634fafe2cd9a5854f", // Exercise (health space)
  "ace998708d25f56dbc8e72a784526a11", // Muscle group
  "ed834cda5168124075774c543866e81d", // Exercise equipment
  "ef193dcb3282afebe466b46b8441c479", // Training category
  "e4b2fe991b8908e4d4cf066e315e4ec8", // Exercise type
  "87ceec69a50da77076f546ad975dd06a", // Body system
  "0ef12e0df2e4478c96cdfd3901109b16", // Difficulty
]);

// Case 4/5: wrong-batch / wrong-logic — entity names to target (from Excel)
// Fill this with the names of the entities in the affected batch
const WRONG_BATCH_NAMES = new Set<string>([
  // e.g. "Side plank", "Rope climb", "Ball slam"
]);

// ── GEO META-IDs ──────────────────────────────────────────────────────────────
const SCHEMA_TYPE_ID   = "e7d737c536764c609fa16aa64a8c90ad"; // "Type" meta-type
const PROPERTY_TYPE_ID = "808a04ceb21c4d888ad12e240613e5ca"; // "Property" meta-type

// ── GEO API HELPER ────────────────────────────────────────────────────────────
async function gql(query: string): Promise<any> {
  const res = await fetch(GEO_API, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ query }),
  });
  return (await res.json() as any).data;
}

// ── LOAD ALL ENTITIES FROM A SPACE (paginated) ────────────────────────────────
interface GeoEntity {
  id: string;
  name: string | null;
  typeIds: string[];
}

async function loadSpace(spaceId: string, label: string): Promise<GeoEntity[]> {
  const all: GeoEntity[] = [];
  const PAGE = 1000;
  let offset = 0;
  process.stdout.write(`Loading ${label}...`);
  while (true) {
    const data = await gql(`{
      entities(spaceId: "${spaceId}", first: ${PAGE}, offset: ${offset}) {
        id name typeIds
      }
    }`);
    const page: GeoEntity[] = data?.entities ?? [];
    all.push(...page);
    offset += PAGE;
    if (page.length < PAGE) break;
  }
  console.log(` ${all.length} entities`);
  return all;
}

// ── OWNERSHIP CHECK ───────────────────────────────────────────────────────────
// Core rule: if entity ID exists in root space → GEO owns it → never delete
function weOwn(entity: GeoEntity, rootIds: Set<string>): boolean {
  return !rootIds.has(entity.id);
}

// ── LOAD INCOMING RELATIONS (who points to each entity) ───────────────────────
async function loadIncomingRelations(entityIds: string[]): Promise<Map<string, Set<string>>> {
  const incomingMap = new Map<string, Set<string>>();
  const CHUNK = 50;
  for (let i = 0; i < entityIds.length; i += CHUNK) {
    const chunk = entityIds.slice(i, i + CHUNK);
    const idsArg = chunk.map(id => `"${id}"`).join(", ");
    const data = await gql(`{
      relations(filter: { toEntityId: { in: [${idsArg}] } }) {
        fromEntity { id }
        toEntity   { id }
      }
    }`);
    for (const r of data?.relations ?? []) {
      const f = r.fromEntity?.id;
      const t = r.toEntity?.id;
      if (!f || !t) continue;
      if (!incomingMap.has(t)) incomingMap.set(t, new Set());
      incomingMap.get(t)!.add(f);
    }
  }
  return incomingMap;
}

// ── ORPHAN CHECK ──────────────────────────────────────────────────────────────
// A supporting entity is orphaned if every entity pointing to it is in the delete set
async function findOrphans(
  deleteSet: Set<string>,
  targetEntities: GeoEntity[],
  rootIds: Set<string>
): Promise<GeoEntity[]> {
  // Collect supporting entities: in target space, not in delete set, not root-owned
  const supporting = targetEntities.filter(e =>
    !deleteSet.has(e.id) && weOwn(e, rootIds)
  );
  if (supporting.length === 0) return [];

  console.log(`  Checking ${supporting.length} potential supporting entities for orphans...`);
  const incomingMap = await loadIncomingRelations(supporting.map(e => e.id));

  return supporting.filter(e => {
    const incomers = incomingMap.get(e.id) ?? new Set();
    // Orphaned if: no incomers at all, OR every incomer is in the delete set
    return incomers.size === 0 || [...incomers].every(id => deleteSet.has(id));
  });
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  const privateKey = process.env.PRIVATE_KEY as `0x${string}`;
  if (!privateKey?.startsWith("0x")) throw new Error("Invalid PRIVATE_KEY in .env");

  // ── 1. LOAD SPACES ─────────────────────────────────────────────────────────
  const targetEntities = await loadSpace(TARGET_SPACE, "personal space");
  const rootEntities   = await loadSpace(ROOT_SPACE,   "root space    ");
  const rootIds = new Set(rootEntities.map(e => e.id));

  // ── 2. IDENTIFY ENTITIES TO DELETE BY MODE ─────────────────────────────────
  console.log(`\nMode: ${MODE}`);
  let primaryTargets: GeoEntity[] = [];

  if (MODE === "wrong-space") {
    // Case 1: delete everything we own — entity ID not in root space
    // Excludes: locally-created Types and Properties are caught here too
    primaryTargets = targetEntities.filter(e => weOwn(e, rootIds));
    console.log(`  Rule: delete all entities we own (not in root space)`);

  } else if (MODE === "wrong-type") {
    // Case 2: delete instances of wrong typeIds
    // Also catch locally-created Type and Property entities
    primaryTargets = targetEntities.filter(e => {
      if (!weOwn(e, rootIds)) return false;
      // Wrong type instances
      if (e.typeIds.some(tid => WRONG_TYPE_IDS.has(tid))) return true;
      // Locally-created Type entities (created by Graph.createType)
      if (e.typeIds.includes(SCHEMA_TYPE_ID)) return true;
      // Locally-created Property entities (created by Graph.createType properties)
      if (e.typeIds.includes(PROPERTY_TYPE_ID)) return true;
      return false;
    });
    console.log(`  Rule: instances of wrong types + locally-created Type/Property entities`);

  } else if (MODE === "wrong-batch" || MODE === "wrong-logic") {
    // Case 4/5: delete entities by name from the affected batch
    if (WRONG_BATCH_NAMES.size === 0) {
      throw new Error("WRONG_BATCH_NAMES is empty — fill in the entity names to delete");
    }
    primaryTargets = targetEntities.filter(e =>
      weOwn(e, rootIds) &&
      WRONG_BATCH_NAMES.has((e.name ?? "").trim())
    );
    console.log(`  Rule: entities matching batch names (${WRONG_BATCH_NAMES.size} names provided)`);
  }

  if (primaryTargets.length === 0) {
    console.log("\n✓ Nothing to delete — space already clean.");
    return;
  }

  console.log(`\nPrimary targets (${primaryTargets.length}):`);
  primaryTargets.forEach(e => console.log(`  ${e.id} | ${e.name ?? "(no name)"}`));

  // ── 3. ORPHAN CHECK (Cases 4/5 — batch deletes) ────────────────────────────
  // For wrong-space and wrong-type we delete everything we own so orphan check
  // is not needed. For batch modes, check what supporting entities become orphaned.
  let orphans: GeoEntity[] = [];

  if (MODE === "wrong-batch" || MODE === "wrong-logic") {
    console.log("\nChecking for orphaned supporting entities...");
    const deleteSet = new Set(primaryTargets.map(e => e.id));
    orphans = await findOrphans(deleteSet, targetEntities, rootIds);

    if (orphans.length > 0) {
      console.log(`  Orphaned supporting entities (${orphans.length}):`);
      orphans.forEach(e => console.log(`    ${e.id} | ${e.name ?? "(no name)"}`));
    } else {
      console.log("  No orphaned supporting entities found.");
    }
  }

  // ── 4. BUILD FINAL DELETE LIST ─────────────────────────────────────────────
  const allToDelete = [...primaryTargets, ...orphans];
  console.log(`\nTotal to delete: ${allToDelete.length}` +
    (orphans.length > 0 ? ` (${primaryTargets.length} primary + ${orphans.length} orphans)` : ""));

  // ── 5. GENERATE DELETE OPS ─────────────────────────────────────────────────
  console.log("Generating delete ops...");
  const allOps: any[] = [];
  for (const e of allToDelete) {
    const result = await Graph.deleteEntity({ id: e.id, spaceId: TARGET_SPACE, network: "TESTNET" });
    allOps.push(...result.ops);
  }

  if (allOps.length === 0) {
    console.log("No ops generated — exiting.");
    return;
  }

  // ── 6. PUBLISH ─────────────────────────────────────────────────────────────
  console.log(`Total ops: ${allOps.length}`);
  console.log("Publishing to IPFS...");

  const editLabel = `Delete edit [${MODE}] — ${allToDelete.length} entities`;

  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: editLabel,
    spaceId: TARGET_SPACE,
    ops: allOps,
    author: TARGET_SPACE,
    network: "TESTNET",
  });

  console.log("Getting wallet...");
  const walletClient = await getSmartAccountWalletClient({ privateKey });

  console.log("Sending transaction...");
  const txHash = await walletClient.sendTransaction({ to, data: calldata });

  console.log(`\n✓ Success!`);
  console.log(`  Mode     : ${MODE}`);
  console.log(`  Deleted  : ${allToDelete.length} entities`);
  console.log(`  Edit ID  : ${editId}`);
  console.log(`  Tx Hash  : ${txHash}`);
  console.log(`  View     : https://geobrowser.io/space/${TARGET_SPACE}`);
}

main().catch(err => {
  console.error("✗ Error:", err.message ?? err);
  process.exit(1);
});
