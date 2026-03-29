import "dotenv/config";
import XLSX from "xlsx";
import {
  Graph,
  personalSpace,
  getSmartAccountWalletClient,
} from "@geoprotocol/geo-sdk";

// ── CONFIGURATION ─────────────────────────────────────────────────────────────
const TARGET_SPACE = "0c747ad3af58ed6b27221a256498068e";
const ROOT_SPACE   = "a19c345ab9866679b001d7d2138d88a1";
const GEO_API      = "https://testnet-api.geobrowser.io/graphql";
const EXCEL_PATH   = "/Users/levan/Desktop/GEO SDK/Exercises Catalog.xlsx";
const BATCH_OFFSET = 100;
const BATCH_LIMIT  = 5;

// ── EXERCISE TYPE (from health space) ────────────────────────────────────────
const EXERCISE_TYPE_ID = "1362f6523665771634fafe2cd9a5854f";

// ── GEO META-IDs ─────────────────────────────────────────────────────────────
const SCHEMA_TYPE_ID         = "e7d737c536764c609fa16aa64a8c90ad";
const PROPERTIES_RELATION_ID = "01412f8381894ab1836565c7fd358cc1";
const RELATION_VALUE_TYPE_ID = "9eea393f17dd4971a62ea603e8bfec20";
const NAME_PROPERTY_ID       = "a126ca530c8e48d5b88882c734c38935";
const DESCRIPTION_PROPERTY_ID= "9b1f76ff9711404c861e59dc3fa7d037";

// ── EXCEL COLUMN → GEO PROPERTY NAME (where they differ) ─────────────────────
// Key = Excel column name, Value = GEO property name
const COLUMN_TO_PROP: Record<string, string> = {
  "Related topics": "Related entities",
  "Body systems ":  "Body systems",   // Excel has a trailing space
};

function colForProp(propName: string): string {
  for (const [col, geo] of Object.entries(COLUMN_TO_PROP)) {
    if (geo === propName) return col;
  }
  return propName;
}

// ── INTERFACES ────────────────────────────────────────────────────────────────
interface GeoEntity {
  id: string;
  name: string | null;
  description: string | null;
  spaceIds: string[];
  typeIds: string[];
}

interface GeoProperty {
  id: string;
  name: string;
  expectedTypeId: string | null;
  expectedTypeName: string | null;
}

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
async function loadSpace(spaceId: string, label: string): Promise<GeoEntity[]> {
  const all: GeoEntity[] = [];
  const PAGE = 1000;
  let offset = 0;
  process.stdout.write(`Loading ${label}...`);
  while (true) {
    const data = await gql(`{
      entities(spaceId: "${spaceId}", first: ${PAGE}, offset: ${offset}) {
        id name description spaceIds typeIds
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

// ── BUILD INDEXES ─────────────────────────────────────────────────────────────
function byName(entities: GeoEntity[]): Map<string, GeoEntity[]> {
  const map = new Map<string, GeoEntity[]>();
  for (const e of entities) {
    if (!e.name) continue;
    const key = e.name.toLowerCase().trim();
    if (!map.has(key)) map.set(key, []);
    map.get(key)!.push(e);
  }
  return map;
}

function byId(entities: GeoEntity[]): Map<string, GeoEntity> {
  return new Map(entities.map(e => [e.id, e]));
}

// ── GET PROPERTIES OF A TYPE (with expected entity type per relation property) ─
const propCache = new Map<string, GeoProperty[]>();

async function getTypeProperties(typeId: string): Promise<GeoProperty[]> {
  if (propCache.has(typeId)) return propCache.get(typeId)!;

  const data = await gql(`{
    entity(id: "${typeId}") {
      relationsList {
        typeId
        toEntity { id name spaceIds
          relationsList(filter: { typeId: { in: ["${RELATION_VALUE_TYPE_ID}"] } }) {
            toEntity { id name }
          }
        }
      }
    }
  }`);

  const props: GeoProperty[] = [];
  for (const rel of data?.entity?.relationsList ?? []) {
    if (rel.typeId !== PROPERTIES_RELATION_ID) continue;
    const pe = rel.toEntity;
    if (!pe?.name) continue;
    const valueTypeRel = pe.relationsList?.find((r: any) => r.toEntity?.name);
    props.push({
      id: pe.id,
      name: pe.name,
      expectedTypeId:   valueTypeRel?.toEntity?.id   ?? null,
      expectedTypeName: valueTypeRel?.toEntity?.name ?? null,
    });
  }

  propCache.set(typeId, props);
  return props;
}

// ── FIND ENTITY BY NAME + TYPE IN TARGET/ROOT SPACE ───────────────────────────
function findEntity(
  name: string,
  typeId: string,
  targetIdx: Map<string, GeoEntity[]>,
  rootIdx: Map<string, GeoEntity[]>
): GeoEntity | null {
  const key = name.toLowerCase().trim();
  for (const idx of [targetIdx, rootIdx]) {
    const match = (idx.get(key) ?? []).find(e => e.typeIds.includes(typeId));
    if (match) return match;
  }
  return null;
}

// ── FIND ENTITY BY NAME ONLY (for properties with meta-type as expected type) ─
function findEntityByName(
  name: string,
  targetIdx: Map<string, GeoEntity[]>,
  rootIdx: Map<string, GeoEntity[]>
): GeoEntity | null {
  const key = name.toLowerCase().trim();
  for (const idx of [targetIdx, rootIdx]) {
    const candidates = idx.get(key) ?? [];
    if (candidates.length > 0) return candidates[0];
  }
  return null;
}

// ── UNIVERSAL FIND-OR-CREATE FOR RELATION TARGETS ────────────────────────────
// idMap key: "typeId::name" to avoid collisions across entity types
function findOrCreate(
  name: string,
  prop: GeoProperty,
  targetIdx: Map<string, GeoEntity[]>,
  rootIdx: Map<string, GeoEntity[]>,
  idMap: Map<string, string>,
  allOps: any[]
): string | null {
  const trimmed = name.trim();
  if (!trimmed) return null;

  const mapKey = `${prop.expectedTypeId}::${trimmed.toLowerCase()}`;
  if (idMap.has(mapKey)) return idMap.get(mapKey)!;

  // If expected type is the meta-type "Type" — search by name only, never auto-create
  // (can't know what type to create; these should link to existing entities)
  if (prop.expectedTypeId === SCHEMA_TYPE_ID) {
    const existing = findEntityByName(trimmed, targetIdx, rootIdx);
    if (existing) {
      idMap.set(mapKey, existing.id);
      return existing.id;
    }
    console.warn(`    ⚠ "${trimmed}" not found for "${prop.name}" (generic relation — skipping, cannot auto-create)`);
    return null;
  }

  // Normal case: find by name + expected type
  const existing = findEntity(trimmed, prop.expectedTypeId!, targetIdx, rootIdx);
  if (existing) {
    idMap.set(mapKey, existing.id);
    return existing.id;
  }

  // Not found → create minimal entity (name only, leave other properties empty)
  console.log(`    + Creating "${trimmed}" as ${prop.expectedTypeName ?? "entity"} (not found in target/root space)`);
  const e = Graph.createEntity({
    name: trimmed,
    types: [prop.expectedTypeId!],
    values: [{ property: NAME_PROPERTY_ID, type: "text", value: trimmed }],
  });
  allOps.push(...e.ops);
  idMap.set(mapKey, e.id);

  // Add to targetIdx so subsequent exercises in this batch find it
  const key = trimmed.toLowerCase();
  if (!targetIdx.has(key)) targetIdx.set(key, []);
  targetIdx.get(key)!.push({
    id: e.id, name: trimmed, description: null,
    spaceIds: [TARGET_SPACE], typeIds: [prop.expectedTypeId!],
  });

  return e.id;
}

// ── GET CURRENT DATA OF AN EXISTING ENTITY (for patch mode) ──────────────────
interface EntityCurrentData {
  filledTextProps: Set<string>;
  filledRelations: Map<string, Set<string>>;
}

async function getEntityCurrentData(entityId: string): Promise<EntityCurrentData> {
  const data = await gql(`{
    entity(id: "${entityId}") {
      valuesList  { propertyId text }
      relationsList { typeId toEntity { id } }
    }
  }`);
  const filledTextProps = new Set<string>();
  for (const v of data?.entity?.valuesList ?? []) {
    if (v.text) filledTextProps.add(v.propertyId);
  }
  const filledRelations = new Map<string, Set<string>>();
  for (const r of data?.entity?.relationsList ?? []) {
    if (!filledRelations.has(r.typeId)) filledRelations.set(r.typeId, new Set());
    if (r.toEntity?.id) filledRelations.get(r.typeId)!.add(r.toEntity.id);
  }
  return { filledTextProps, filledRelations };
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  const privateKey = process.env.PRIVATE_KEY as `0x${string}`;
  if (!privateKey?.startsWith("0x")) throw new Error("Invalid PRIVATE_KEY in .env");

  // ── 1. LOAD BOTH SPACES ────────────────────────────────────────────────────
  const targetEntities = await loadSpace(TARGET_SPACE, "target space");
  const rootEntities   = await loadSpace(ROOT_SPACE,   "root space  ");

  const targetIdx = byName(targetEntities);
  const rootIdx   = byName(rootEntities);

  // ── 2. LOAD EXERCISE TYPE PROPERTIES (dynamic — no hardcoding) ────────────
  console.log("\nResolving Exercise type properties...");
  const exerciseProps = await getTypeProperties(EXERCISE_TYPE_ID);

  if (exerciseProps.length === 0) {
    throw new Error("Exercise type has no properties — cannot continue.");
  }

  console.log(`  Found ${exerciseProps.length} properties:`);
  for (const p of exerciseProps) {
    const kind = p.expectedTypeId ? `RELATION → ${p.expectedTypeName}` : "TEXT";
    console.log(`    ${p.name.padEnd(22)} [${kind}]`);
  }

  // ── 3. READ EXCEL ──────────────────────────────────────────────────────────
  console.log("\nReading Excel...");
  const wb = XLSX.readFile(EXCEL_PATH);
  const exercises = XLSX.utils.sheet_to_json<any>(wb.Sheets["Exercises"]);
  const batch = exercises.slice(BATCH_OFFSET, BATCH_OFFSET + BATCH_LIMIT);
  console.log(`  Processing exercises ${BATCH_OFFSET + 1}–${BATCH_OFFSET + batch.length} of ${exercises.length}`);

  // ── 4. PROCESS EXERCISES ──────────────────────────────────────────────────
  const allOps: any[] = [];
  const idMap = new Map<string, string>(); // "typeId::name" → entity id
  let newCount = 0;
  let patchedCount = 0;

  for (let i = 0; i < batch.length; i++) {
    const ex = batch[i];
    const name = (ex["Entity Name"] as string)?.trim();
    if (!name) continue;

    console.log(`\n[${BATCH_OFFSET + i + 1}] ${name}`);

    // Check if exercise already exists
    const existing = findEntity(name, EXERCISE_TYPE_ID, targetIdx, rootIdx);

    if (existing) {
      // ── PATCH: fill in only missing properties ─────────────────────────────
      console.log(`  → exists (${existing.id}) — patching missing properties`);
      const current = await getEntityCurrentData(existing.id);
      const missingText: any[] = [];

      for (const geoprop of exerciseProps) {
        const colName = colForProp(geoprop.name);
        const rawVal  = (ex[colName] ?? "").toString().trim();
        if (!rawVal) continue;

        if (geoprop.expectedTypeId === null) {
          // Text property — add if not already filled
          if (!current.filledTextProps.has(geoprop.id)) {
            missingText.push({ property: geoprop.id, type: "text", value: rawVal });
          }
        } else {
          // Relation property — add only missing links
          const existingLinks = current.filledRelations.get(geoprop.id) ?? new Set();
          for (const val of splitVal(rawVal)) {
            const toId = findOrCreate(val, geoprop, targetIdx, rootIdx, idMap, allOps);
            if (!toId || existingLinks.has(toId)) continue;
            const r = Graph.createRelation({ fromEntity: existing.id, toEntity: toId, type: geoprop.id });
            allOps.push(...r.ops);
          }
        }
      }

      if (missingText.length > 0) {
        const update = Graph.updateEntity({ id: existing.id, values: missingText });
        allOps.push(...update.ops);
        patchedCount++;
      }

    } else {
      // ── CREATE: new exercise entity ────────────────────────────────────────
      console.log(`  → new entity`);
      const textValues: any[] = [{ property: NAME_PROPERTY_ID, type: "text", value: name }];

      // Collect text properties first
      for (const geoprop of exerciseProps) {
        if (geoprop.expectedTypeId !== null) continue; // skip relations for now
        const colName = colForProp(geoprop.name);
        const rawVal  = (ex[colName] ?? "").toString().trim();
        if (rawVal) textValues.push({ property: geoprop.id, type: "text", value: rawVal });
      }

      const exercise = Graph.createEntity({
        name,
        types: [EXERCISE_TYPE_ID],
        values: textValues.filter(v => v.value),
      });
      allOps.push(...exercise.ops);

      // Register in index so same-batch exercises can find this as a variation target
      const key = name.toLowerCase();
      if (!targetIdx.has(key)) targetIdx.set(key, []);
      targetIdx.get(key)!.push({
        id: exercise.id, name, description: null,
        spaceIds: [TARGET_SPACE], typeIds: [EXERCISE_TYPE_ID],
      });

      // Now add relation properties
      for (const geoprop of exerciseProps) {
        if (geoprop.expectedTypeId === null) continue; // skip text props
        const colName = colForProp(geoprop.name);
        const rawVal  = (ex[colName] ?? "").toString().trim();
        if (!rawVal) continue;

        for (const val of splitVal(rawVal)) {
          const toId = findOrCreate(val, geoprop, targetIdx, rootIdx, idMap, allOps);
          if (!toId) continue;
          const r = Graph.createRelation({ fromEntity: exercise.id, toEntity: toId, type: geoprop.id });
          allOps.push(...r.ops);
        }
      }

      newCount++;
    }
  }

  console.log(`\nSummary: ${newCount} new | ${patchedCount} patched`);

  if (allOps.length === 0) {
    console.log("✓ Nothing new to upload.");
    return;
  }

  // ── 5. PUBLISH ─────────────────────────────────────────────────────────────
  console.log(`\nTotal ops: ${allOps.length}`);
  console.log("Publishing to IPFS...");

  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: `Exercise Catalog — Batch 3 (exercises ${BATCH_OFFSET + 1}–${BATCH_OFFSET + BATCH_LIMIT})`,
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
  console.log(`  Edit ID  : ${editId}`);
  console.log(`  Tx Hash  : ${txHash}`);
  console.log(`  View     : https://geobrowser.io/space/${TARGET_SPACE}`);
}

// ── UTILS ─────────────────────────────────────────────────────────────────────
function splitVal(val: unknown): string[] {
  if (!val || typeof val !== "string") return [];
  return val.split(",").map(s => s.trim()).filter(Boolean);
}

main().catch(err => {
  console.error("✗ Error:", err.message ?? err);
  process.exit(1);
});
