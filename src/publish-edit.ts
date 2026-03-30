import "dotenv/config";
import * as readline from "readline";
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
const SHEET_NAME   = "Exercises";
const BATCH_OFFSET = 100;
const BATCH_LIMIT  = 5;

// ── ROOT TYPE TO UPLOAD (from health/root space) ──────────────────────────────
const ROOT_TYPE_ID = "1362f6523665771634fafe2cd9a5854f"; // Exercise

// ── GEO META-IDs ─────────────────────────────────────────────────────────────
const SCHEMA_TYPE_ID         = "e7d737c536764c609fa16aa64a8c90ad";
const PROPERTIES_RELATION_ID = "01412f8381894ab1836565c7fd358cc1";
const RELATION_VALUE_TYPE_ID = "9eea393f17dd4971a62ea603e8bfec20";
const VALUE_TYPE_RELATION_ID = "6d29d57849bb4959baf72cc696b1671a"; // data-type pointer: → Relation / Text / Date / etc.
const NAME_PROPERTY_ID       = "a126ca530c8e48d5b88882c734c38935";
const DESCRIPTION_PROPERTY_ID= "9b1f76ff9711404c861e59dc3fa7d037";

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
  isRelation: boolean;        // true even when expectedTypeId is null (untyped relation)
  expectedTypeId: string | null;
  expectedTypeName: string | null;
}

// Resolved mapping: GEO property name → document column name (null = skip)
type PropertyMapping = Map<string, string | null>;

// ── CLI HELPER ────────────────────────────────────────────────────────────────
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

function ask(prompt: string): Promise<string> {
  return new Promise(resolve => rl.question(prompt, answer => resolve(answer.trim())));
}

function closeInput() { rl.close(); }

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

// ── PICK WHICH SPACE'S VERSION OF A TYPE TO USE ───────────────────────────────
async function pickTypeSpace(typeId: string): Promise<string> {
  const data = await gql(`{
    entity(id: "${typeId}") {
      name
      spaceIds
    }
  }`);

  const typeName  = data?.entity?.name ?? typeId;
  const spaceIds: string[] = data?.entity?.spaceIds ?? [];

  if (spaceIds.length === 0) throw new Error(`Type "${typeName}" not found in any space.`);
  if (spaceIds.length === 1) {
    console.log(`\nType "${typeName}" exists in 1 space — using it automatically.`);
    return spaceIds[0];
  }

  // Resolve space names
  const spacesData = await gql(`{
    spaces(filter: { id: { in: [${spaceIds.map(id => `"${id}"`).join(", ")}] } }) {
      id
      page { name }
    }
  }`);
  const spaceNames = new Map<string, string>();
  for (const s of spacesData?.spaces ?? []) {
    spaceNames.set(s.id, s.page?.name ?? s.id);
  }

  console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("  TYPE SPACE SELECTION");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log(`\n  Type "${typeName}" exists in ${spaceIds.length} spaces with different property sets:\n`);
  spaceIds.forEach((id, i) => {
    const label = spaceNames.get(id) ?? id;
    console.log(`    ${i + 1}. ${label.padEnd(20)} (${id})`);
  });
  console.log();

  let choice = -1;
  while (choice < 1 || choice > spaceIds.length) {
    const input = await ask(`  Which space's property set do you want to use? Enter number: `);
    choice = parseInt(input);
    if (isNaN(choice) || choice < 1 || choice > spaceIds.length) {
      console.log(`  Please enter a number between 1 and ${spaceIds.length}.`);
      choice = -1;
    }
  }

  const chosen = spaceIds[choice - 1];
  console.log(`\n  ✅ Using "${spaceNames.get(chosen) ?? chosen}" property set\n`);
  return chosen;
}

// ── GET PROPERTIES OF A TYPE (filtered by space) ──────────────────────────────
const propCache = new Map<string, GeoProperty[]>();

async function getTypeProperties(typeId: string, spaceId: string): Promise<GeoProperty[]> {
  const cacheKey = `${typeId}::${spaceId}`;
  if (propCache.has(cacheKey)) return propCache.get(cacheKey)!;

  const data = await gql(`{
    entity(id: "${typeId}") {
      relationsList {
        typeId
        spaceId
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
    if (rel.spaceId !== spaceId) continue;
    const pe = rel.toEntity;
    if (!pe?.name) continue;
    const subRels: any[] = pe.relationsList ?? [];
    // Data type: the VALUE_TYPE_RELATION_ID sub-relation tells us if this is a Relation, Text, Date, etc.
    const dataTypeRel  = subRels.find((r: any) => r.typeId === VALUE_TYPE_RELATION_ID);
    const isRelation   = dataTypeRel?.toEntity?.name === "Relation";
    // Expected entity type: only present for typed relations
    const valueTypeRel = subRels.find((r: any) => r.typeId === RELATION_VALUE_TYPE_ID);
    props.push({
      id: pe.id,
      name: pe.name,
      isRelation,
      expectedTypeId:   valueTypeRel?.toEntity?.id   ?? null,
      expectedTypeName: valueTypeRel?.toEntity?.name ?? null,
    });
  }

  propCache.set(cacheKey, props);
  return props;
}

// ── INTERACTIVE MAPPING BUILDER ───────────────────────────────────────────────
async function buildMapping(
  geoProps: GeoProperty[],
  docColumns: string[]
): Promise<PropertyMapping> {
  const mapping: PropertyMapping = new Map();
  const usedDocCols = new Set<string>();

  console.log("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("  PROPERTY MAPPING");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

  for (const geoprop of geoProps) {
    const kind = geoprop.isRelation
      ? (geoprop.expectedTypeName ? `RELATION → ${geoprop.expectedTypeName}` : "RELATION (untyped)")
      : "TEXT";

    // Try auto-match: exact name (case-insensitive, trimmed)
    const autoMatch = docColumns.find(
      col => col.trim().toLowerCase() === geoprop.name.toLowerCase()
    );

    if (autoMatch && !usedDocCols.has(autoMatch)) {
      mapping.set(geoprop.name, autoMatch);
      usedDocCols.add(autoMatch);
      console.log(`  ✅ "${geoprop.name}" [${kind}]`);
      console.log(`     → auto-matched to column "${autoMatch}"\n`);
      continue;
    }

    // No auto-match — ask the user
    console.log(`  ⚠️  "${geoprop.name}" [${kind}]`);
    console.log(`     No matching column found.\n`);

    const available = docColumns.filter(col => !usedDocCols.has(col));

    console.log(`     Available document columns:`);
    available.forEach((col, i) => console.log(`       ${i + 1}. "${col}"`));
    console.log(`       ${available.length + 1}. Skip this property (leave unfilled)`);

    let choice = -1;
    while (choice < 1 || choice > available.length + 1) {
      const input = await ask(`\n     Which column maps to "${geoprop.name}"? Enter number: `);
      choice = parseInt(input);
      if (isNaN(choice) || choice < 1 || choice > available.length + 1) {
        console.log(`     Please enter a number between 1 and ${available.length + 1}.`);
        choice = -1;
      }
    }

    if (choice === available.length + 1) {
      mapping.set(geoprop.name, null);
      console.log(`     → Skipped\n`);
    } else {
      const chosen = available[choice - 1];
      mapping.set(geoprop.name, chosen);
      usedDocCols.add(chosen);
      console.log(`     → Mapped to "${chosen}"\n`);
    }
  }

  // Warn about document columns not mapped to any GEO property
  const unmappedDocCols = docColumns.filter(col => !usedDocCols.has(col));
  if (unmappedDocCols.length > 0) {
    console.log(`  ℹ️  Document columns with no GEO property match (will be ignored):`);
    unmappedDocCols.forEach(col => console.log(`     - "${col}"`));
    console.log();
  }

  // Full mapping summary
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("  MAPPING SUMMARY");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  for (const [geoProp, docCol] of mapping) {
    const status = docCol ? `→ "${docCol}"` : "→ skipped";
    console.log(`  ${geoProp.padEnd(25)} ${status}`);
  }
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");

  // Final confirmation
  const confirm = await ask("  Proceed with this mapping? (y/n): ");
  if (confirm.toLowerCase() !== "y") {
    console.log("\n  Upload cancelled.");
    closeInput();
    process.exit(0);
  }

  console.log();
  return mapping;
}

// ── FIND ENTITY BY NAME + TYPE ────────────────────────────────────────────────
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

  // Untyped relation or generic meta-type: search by name only, never auto-create
  if (!prop.expectedTypeId || prop.expectedTypeId === SCHEMA_TYPE_ID) {
    const existing = findEntityByName(trimmed, targetIdx, rootIdx);
    if (existing) { idMap.set(mapKey, existing.id); return existing.id; }
    console.warn(`    ⚠ "${trimmed}" not found for "${prop.name}" (untyped relation — skipping)`);
    return null;
  }

  // Typed relation: find by name + expected type
  const existing = findEntity(trimmed, prop.expectedTypeId!, targetIdx, rootIdx);
  if (existing) { idMap.set(mapKey, existing.id); return existing.id; }

  // Not found → create minimal entity (name only)
  console.log(`    + Creating "${trimmed}" as ${prop.expectedTypeName ?? "entity"}`);
  const e = Graph.createEntity({
    name: trimmed,
    types: [prop.expectedTypeId!],
    values: [{ property: NAME_PROPERTY_ID, type: "text", value: trimmed }],
  });
  allOps.push(...e.ops);
  idMap.set(mapKey, e.id);

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

  // ── 2. PICK SPACE VERSION OF TYPE + LOAD PROPERTIES ──────────────────────
  const typeSpaceId = await pickTypeSpace(ROOT_TYPE_ID);
  console.log("Resolving type properties from GEO...");
  const geoProps = await getTypeProperties(ROOT_TYPE_ID, typeSpaceId);
  if (geoProps.length === 0) throw new Error("Type has no properties defined in GEO for the selected space.");
  console.log(`  Found ${geoProps.length} properties on type`);

  // ── 3. READ DOCUMENT ───────────────────────────────────────────────────────
  console.log("\nReading document...");
  const wb = XLSX.readFile(EXCEL_PATH);
  const rows = XLSX.utils.sheet_to_json<any>(wb.Sheets[SHEET_NAME]);
  const batch = rows.slice(BATCH_OFFSET, BATCH_OFFSET + BATCH_LIMIT);

  // Get document columns (exclude entity name and type columns)
  const docColumns = Object.keys(rows[0] ?? {}).filter(col =>
    col !== "Entity Name" && col !== "Type"
  );
  console.log(`  ${rows.length} rows | ${docColumns.length} property columns | batch: ${BATCH_OFFSET + 1}–${BATCH_OFFSET + batch.length}`);

  // ── 4. INTERACTIVE MAPPING ─────────────────────────────────────────────────
  const mapping = await buildMapping(geoProps, docColumns);

  // ── 5. PROCESS ENTITIES ────────────────────────────────────────────────────
  console.log(`Processing ${batch.length} entities...\n`);

  const allOps: any[] = [];
  const idMap = new Map<string, string>();
  let newCount = 0;
  let patchedCount = 0;

  for (let i = 0; i < batch.length; i++) {
    const row = batch[i];
    const name = (row["Entity Name"] as string)?.trim();
    if (!name) continue;

    console.log(`[${BATCH_OFFSET + i + 1}] ${name}`);

    const existing = findEntity(name, ROOT_TYPE_ID, targetIdx, rootIdx);

    if (existing) {
      // ── PATCH ──────────────────────────────────────────────────────────────
      console.log(`  → exists — patching missing properties`);
      const current = await getEntityCurrentData(existing.id);
      const missingText: any[] = [];

      for (const geoprop of geoProps) {
        const docCol = mapping.get(geoprop.name);
        if (!docCol) continue;
        const rawVal = (row[docCol] ?? "").toString().trim();
        if (!rawVal) continue;

        if (!geoprop.isRelation) {
          if (!current.filledTextProps.has(geoprop.id))
            missingText.push({ property: geoprop.id, type: "text", value: rawVal });
        } else {
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
      // ── CREATE ─────────────────────────────────────────────────────────────
      console.log(`  → new entity`);
      const textValues: any[] = [{ property: NAME_PROPERTY_ID, type: "text", value: name }];

      for (const geoprop of geoProps) {
        if (geoprop.isRelation) continue;
        const docCol = mapping.get(geoprop.name);
        if (!docCol) continue;
        const rawVal = (row[docCol] ?? "").toString().trim();
        if (rawVal) textValues.push({ property: geoprop.id, type: "text", value: rawVal });
      }

      const entity = Graph.createEntity({
        name,
        types: [ROOT_TYPE_ID],
        values: textValues.filter(v => v.value),
      });
      allOps.push(...entity.ops);

      const key = name.toLowerCase();
      if (!targetIdx.has(key)) targetIdx.set(key, []);
      targetIdx.get(key)!.push({
        id: entity.id, name, description: null,
        spaceIds: [TARGET_SPACE], typeIds: [ROOT_TYPE_ID],
      });

      for (const geoprop of geoProps) {
        if (!geoprop.isRelation) continue;
        const docCol = mapping.get(geoprop.name);
        if (!docCol) continue;
        const rawVal = (row[docCol] ?? "").toString().trim();
        if (!rawVal) continue;

        for (const val of splitVal(rawVal)) {
          const toId = findOrCreate(val, geoprop, targetIdx, rootIdx, idMap, allOps);
          if (!toId) continue;
          const r = Graph.createRelation({ fromEntity: entity.id, toEntity: toId, type: geoprop.id });
          allOps.push(...r.ops);
        }
      }

      newCount++;
    }
  }

  closeInput();
  console.log(`\nSummary: ${newCount} new | ${patchedCount} patched`);

  if (allOps.length === 0) {
    console.log("✓ Nothing new to upload.");
    return;
  }

  // ── 6. PUBLISH ─────────────────────────────────────────────────────────────
  console.log(`\nTotal ops: ${allOps.length}`);
  console.log("Publishing to IPFS...");

  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: `Batch ${BATCH_OFFSET + 1}–${BATCH_OFFSET + BATCH_LIMIT}`,
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
  closeInput();
  console.error("✗ Error:", err.message ?? err);
  process.exit(1);
});
