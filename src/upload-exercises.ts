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
const EXCEL_PATH   = "/Users/levan/Desktop/My best friend/Exercises Catalog (1).xlsx";
const BATCH_OFFSET = 50;
const BATCH_LIMIT  = 50;

// ── EXERCISE TYPE (from health space) ────────────────────────────────────────
const EXERCISE_TYPE_ID = "1362f6523665771634fafe2cd9a5854f";

// ── GEO META-IDs (fixed, from Root Space) ────────────────────────────────────
const SCHEMA_TYPE_ID               = "e7d737c536764c609fa16aa64a8c90ad"; // the "Type" meta-type
const PROPERTIES_RELATION_ID       = "01412f8381894ab1836565c7fd358cc1"; // type → its properties
const RELATION_VALUE_TYPE_ID       = "9eea393f17dd4971a62ea603e8bfec20"; // property → expected entity type
const NAME_PROPERTY_ID             = "a126ca530c8e48d5b88882c734c38935";
const DESCRIPTION_PROPERTY_ID      = "9b1f76ff9711404c861e59dc3fa7d037";

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
  spaceIds: string[];
  expectedTypeId: string | null;   // what entity type this relation property expects
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

// ── PICK BEST DUPLICATE (prefer entity with description) ─────────────────────
function pickBest(candidates: GeoEntity[]): GeoEntity | null {
  if (!candidates.length) return null;
  return candidates.find(e => e.description) ?? candidates[0];
}

// ── FIND TYPE: target space first → root space ────────────────────────────────
// A type is an entity whose typeIds includes SCHEMA_TYPE_ID
function findType(
  name: string,
  targetIdx: Map<string, GeoEntity[]>,
  rootIdx: Map<string, GeoEntity[]>
): GeoEntity | null {
  const key = name.toLowerCase().trim();
  for (const idx of [targetIdx, rootIdx]) {
    const match = (idx.get(key) ?? []).find(e => e.typeIds.includes(SCHEMA_TYPE_ID));
    if (match) return match;
  }
  console.warn(`  ⚠ Type not found: "${name}" — skipping`);
  return null;
}

// ── GET PROPERTIES OF A TYPE ──────────────────────────────────────────────────
// Queries the type entity's PROPERTIES relations and each property's expected entity type
const propCache = new Map<string, GeoProperty[]>();

async function getTypeProperties(typeId: string, idIdx: Map<string, GeoEntity>): Promise<GeoProperty[]> {
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

  const relations: any[] = data?.entity?.relationsList ?? [];
  const props: GeoProperty[] = [];

  for (const rel of relations) {
    if (rel.typeId !== PROPERTIES_RELATION_ID) continue;
    const pe = rel.toEntity;
    if (!pe?.name) continue;

    // Find the expected entity type for this property (if it's a relation property)
    const valueTypeRel = pe.relationsList?.find((r: any) => r.toEntity?.name);
    const expectedTypeId   = valueTypeRel?.toEntity?.id ?? null;
    const expectedTypeName = valueTypeRel?.toEntity?.name ?? null;

    props.push({ id: pe.id, name: pe.name, spaceIds: pe.spaceIds, expectedTypeId, expectedTypeName });
  }

  propCache.set(typeId, props);
  return props;
}

// ── GET CURRENT DATA OF AN EXISTING ENTITY ───────────────────────────────────
// Returns which text propertyIds are already filled, and which relation typeIds exist
interface EntityCurrentData {
  filledTextProps: Set<string>;          // propertyIds that already have a text value
  filledRelations: Map<string, Set<string>>; // typeId → set of toEntity IDs already linked
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

// ── FIND ENTITY FOR RELATION ──────────────────────────────────────────────────
// Matches by: name + expected type + space must be target or root
function findEntityForRelation(
  name: string,
  expectedTypeId: string,
  targetIdx: Map<string, GeoEntity[]>,
  rootIdx: Map<string, GeoEntity[]>
): GeoEntity | null {
  const key = name.toLowerCase().trim();
  for (const [idx, spaceId] of [[targetIdx, TARGET_SPACE], [rootIdx, ROOT_SPACE]] as [Map<string, GeoEntity[]>, string][]) {
    const match = (idx.get(key) ?? []).find(e =>
      e.typeIds.includes(expectedTypeId) &&
      e.spaceIds.includes(spaceId)
    );
    if (match) return match;
  }
  return null;
}

// ── MAIN ──────────────────────────────────────────────────────────────────────
async function main() {
  const privateKey = process.env.PRIVATE_KEY as `0x${string}`;
  if (!privateKey?.startsWith("0x")) throw new Error("Invalid PRIVATE_KEY in .env");

  // ── 1. LOAD BOTH SPACES ────────────────────────────────────────────────────
  const targetEntities = await loadSpace(TARGET_SPACE, "target space");
  const rootEntities   = await loadSpace(ROOT_SPACE,   "root space  ");

  const targetIdx  = byName(targetEntities);
  const rootIdx    = byName(rootEntities);
  const targetIdIdx = byId(targetEntities);
  const rootIdIdx   = byId(rootEntities);
  const allIdIdx    = new Map([...targetIdIdx, ...rootIdIdx]);

  // ── 2. RESOLVE TYPES (never create) ───────────────────────────────────────
  console.log("\nResolving types...");

  // Exercise type is pinned to the health space type ID — no lookup needed
  const exerciseType = { id: EXERCISE_TYPE_ID } as GeoEntity;

  const bodySystemType  = findType("Body system",       targetIdx, rootIdx);
  const difficultyType  = findType("Difficulty level",  targetIdx, rootIdx) ??
                          findType("Difficulty",         targetIdx, rootIdx);
  const exTypeType      = findType("Exercise type",     targetIdx, rootIdx);
  const trainingCatType = findType("Training category", targetIdx, rootIdx);
  const relatedTopicType= findType("Related topic",     targetIdx, rootIdx) ??
                          findType("Topic",              targetIdx, rootIdx);
  const variationType   = findType("Exercise variation",targetIdx, rootIdx) ??
                          findType("Exercise",           targetIdx, rootIdx);

  console.log(`  Exercise type      : ${exerciseType.id} (health space — pinned)`);
  if (bodySystemType)  console.log(`  Body system type   : ${bodySystemType.id}`);
  if (difficultyType)  console.log(`  Difficulty type    : ${difficultyType.id}`);
  if (exTypeType)      console.log(`  Exercise type type : ${exTypeType.id}`);
  if (trainingCatType) console.log(`  Training cat type  : ${trainingCatType.id}`);
  if (relatedTopicType)console.log(`  Related topic type : ${relatedTopicType.id}`);
  if (variationType)   console.log(`  Variation type     : ${variationType.id}`);

  // ── 3. RESOLVE PROPERTIES FROM THE EXERCISE TYPE (never create) ───────────
  console.log("\nResolving Exercise type properties...");
  const exerciseProps = await getTypeProperties(exerciseType.id, allIdIdx);

  if (exerciseProps.length === 0) {
    console.warn("  ⚠ Exercise type has no properties defined. Text fields will use well-known IDs only.");
  } else {
    console.log(`  Found ${exerciseProps.length} properties: ${exerciseProps.map(p => p.name).join(", ")}`);
  }

  // Helper: find property by column name (case-insensitive)
  function prop(colName: string): string | null {
    const p = exerciseProps.find(p => p.name.toLowerCase().trim() === colName.toLowerCase().trim());
    if (!p) {
      // Fall back to well-known IDs for Name and Description
      if (colName.toLowerCase() === "name")        return NAME_PROPERTY_ID;
      if (colName.toLowerCase() === "description") return DESCRIPTION_PROPERTY_ID;
      console.warn(`  ⚠ Property "${colName}" not found on Exercise type — column will be skipped`);
      return null;
    }
    return p.id;
  }

  // ── 4. READ EXCEL ──────────────────────────────────────────────────────────
  console.log("\nReading Excel...");
  const wb = XLSX.readFile(EXCEL_PATH);

  const exercises        = XLSX.utils.sheet_to_json<any>(wb.Sheets["Exercises"]);
  const bodySystemRows   = XLSX.utils.sheet_to_json<any>(wb.Sheets["Body System"]);
  const difficultyRows   = XLSX.utils.sheet_to_json<any>(wb.Sheets["Difficulty"]);
  const exerciseTypeRows = XLSX.utils.sheet_to_json<any>(wb.Sheets["Exercise type"]);
  const relatedTopicRows = XLSX.utils.sheet_to_json<any>(wb.Sheets["Related topics"]);
  const trainingCatRows  = XLSX.utils.sheet_to_json<any>(wb.Sheets["Training Category - minus other"]);
  const variationRows    = XLSX.utils.sheet_to_json<any>(wb.Sheets["Variations"]);

  const batch = exercises.slice(BATCH_OFFSET, BATCH_OFFSET + BATCH_LIMIT);

  const referencedVariations = new Set<string>();
  batch.forEach((ex: any) =>
    splitVal(ex["Variations"]).forEach(v => referencedVariations.add(v))
  );

  // ── 5. PROCESS SUPPORTING ENTITIES ────────────────────────────────────────
  console.log("\nProcessing supporting entities...");

  const allOps: any[] = [];
  const idMap = new Map<string, string>(); // entity name → id
  let reused = 0;
  let created = 0;

  function upsertSupporting(name: string, description: string, typeEntity: GeoEntity | null): void {
    if (!name || idMap.has(name)) return;
    if (!typeEntity) {
      console.warn(`  ⚠ Skipping "${name}" — type not resolved`);
      return;
    }

    // Find existing: name + correct type + target or root space
    const existing = findEntityForRelation(name, typeEntity.id, targetIdx, rootIdx);
    if (existing) {
      idMap.set(name, existing.id);
      reused++;
      return;
    }

    // Create new in target space with the resolved type
    const e = Graph.createEntity({
      name,
      types: [typeEntity.id],
      values: [
        { property: NAME_PROPERTY_ID,        type: "text", value: name },
        { property: DESCRIPTION_PROPERTY_ID, type: "text", value: description },
      ].filter(v => v.value) as any,
    });
    idMap.set(name, e.id);
    allOps.push(...e.ops);
    created++;
  }

  bodySystemRows.forEach((r: any)   => upsertSupporting(r["Entity Name"], r["Description"] ?? "", bodySystemType));
  difficultyRows.forEach((r: any)   => upsertSupporting(r["Entity Name"], r["Description"] ?? "", difficultyType));
  exerciseTypeRows.forEach((r: any) => upsertSupporting(r["Entity Name"], r["Description"] ?? "", exTypeType));
  relatedTopicRows.forEach((r: any) => upsertSupporting(r["Entity Name"], r["Description"] ?? "", relatedTopicType));
  trainingCatRows.filter((r: any) => r["Name"])
    .forEach((r: any) => upsertSupporting(r["Name"], r["Clear definition"] ?? "", trainingCatType));
  variationRows
    .filter((r: any) => r["Entity Name"] && referencedVariations.has(r["Entity Name"]))
    .forEach((r: any) => upsertSupporting(r["Entity Name"], r["Description"] ?? "", variationType));

  console.log(`  Reused: ${reused} | Created: ${created}`);

  // ── 6. PROCESS EXERCISES ──────────────────────────────────────────────────
  console.log(`\nProcessing ${batch.length} exercises...`);

  // Resolve property IDs from the Exercise type (or well-known fallbacks)
  const descPropId          = prop("Description");
  const commonMistakesPropId= prop("Common mistakes");
  const equipmentPropId     = prop("Equipments");
  const formCuesPropId      = prop("Form cues");
  const primaryMusclesPropId= prop("Primary muscles");
  const secondaryMusclesPropId = prop("Secondary muscles");
  const bodySystemPropId    = prop("Body systems");
  const difficultyPropId    = prop("Difficulty");
  const exTypePropId        = prop("Exercise type");
  const relTopicPropId      = prop("Related entities");
  const trainingCatPropId   = prop("Training category");
  const variationsPropId    = prop("Variations");

  let newCount = 0;
  let patchedCount = 0;

  // Process exercises sequentially (needs async for patch queries)
  for (let i = 0; i < batch.length; i++) {
    const ex = batch[i];
    const name = ex["Entity Name"] as string;

    // Check if exercise already exists (name + type + space)
    const existing = findEntityForRelation(name, exerciseType.id, targetIdx, rootIdx);

    if (existing) {
      // ── PATCH: fill in any missing properties ──────────────────────────────
      idMap.set(name, existing.id);
      const current = await getEntityCurrentData(existing.id);

      // Text properties — only add if currently empty
      const missingValues: any[] = [];
      const textCols: [string | null, string][] = [
        [descPropId,             ex["Description"]       ?? ""],
        [commonMistakesPropId,   ex["Common mistakes"]   ?? ""],
        [equipmentPropId,        ex["Equipment"]         ?? ""],
        [formCuesPropId,         ex["Form cues"]         ?? ""],
        [primaryMusclesPropId,   ex["Primary muscles"]   ?? ""],
        [secondaryMusclesPropId, ex["Secondary muscles"] ?? ""],
      ];
      for (const [propId, val] of textCols) {
        if (propId && val && !current.filledTextProps.has(propId)) {
          missingValues.push({ property: propId, type: "text", value: val });
        }
      }
      if (missingValues.length > 0) {
        const update = Graph.updateEntity({ id: existing.id, values: missingValues });
        allOps.push(...update.ops);
        patchedCount++;
      }

      // Relation properties — only add relations that don't already exist
      const relCols: [string[], string | null, string | null][] = [
        [splitVal(ex["Body systems"]),     bodySystemPropId,  bodySystemType?.id   ?? null],
        [splitVal(ex["Difficulty"]),        difficultyPropId,  difficultyType?.id   ?? null],
        [splitVal(ex["Exercise type"]),     exTypePropId,      exTypeType?.id       ?? null],
        [splitVal(ex["Related topics"]),    relTopicPropId,    relatedTopicType?.id ?? null],
        [splitVal(ex["Training category"]), trainingCatPropId, trainingCatType?.id  ?? null],
        [splitVal(ex["Variations"]),        variationsPropId,  variationType?.id    ?? null],
      ];
      for (const [vals, propId, expectedTypeId] of relCols) {
        if (!propId || !expectedTypeId) continue;
        const existingTargets = current.filledRelations.get(propId) ?? new Set();
        for (const val of vals) {
          let toId = idMap.get(val) ?? findEntityForRelation(val, expectedTypeId, targetIdx, rootIdx)?.id;
          if (!toId) { console.warn(`    ⚠ Relation target not found: "${val}"`); continue; }
          if (existingTargets.has(toId)) continue; // already linked
          const r = Graph.createRelation({ fromEntity: existing.id, toEntity: toId, type: propId });
          allOps.push(...r.ops);
        }
      }

    } else {
      // ── CREATE: new entity ─────────────────────────────────────────────────
      const values: any[] = [{ property: NAME_PROPERTY_ID, type: "text", value: name }];
      if (descPropId)             values.push({ property: descPropId,             type: "text", value: ex["Description"]       ?? "" });
      if (commonMistakesPropId)   values.push({ property: commonMistakesPropId,   type: "text", value: ex["Common mistakes"]   ?? "" });
      if (equipmentPropId)        values.push({ property: equipmentPropId,        type: "text", value: ex["Equipment"]         ?? "" });
      if (formCuesPropId)         values.push({ property: formCuesPropId,         type: "text", value: ex["Form cues"]         ?? "" });
      if (primaryMusclesPropId)   values.push({ property: primaryMusclesPropId,   type: "text", value: ex["Primary muscles"]   ?? "" });
      if (secondaryMusclesPropId) values.push({ property: secondaryMusclesPropId, type: "text", value: ex["Secondary muscles"] ?? "" });

      const exercise = Graph.createEntity({
        name,
        types: [exerciseType.id],
        values: values.filter(v => v.value),
      });
      allOps.push(...exercise.ops);
      idMap.set(name, exercise.id);
      newCount++;

      // Relations
      const rel = (vals: string[], propId: string | null, expectedTypeId: string | null) => {
        if (!propId || !expectedTypeId) return;
        for (const val of vals) {
          const toId = idMap.get(val) ?? findEntityForRelation(val, expectedTypeId, targetIdx, rootIdx)?.id;
          if (!toId) { console.warn(`    ⚠ Relation target not found: "${val}"`); continue; }
          const r = Graph.createRelation({ fromEntity: exercise.id, toEntity: toId, type: propId });
          allOps.push(...r.ops);
        }
      };

      rel(splitVal(ex["Body systems"]),      bodySystemPropId,  bodySystemType?.id   ?? null);
      rel(splitVal(ex["Difficulty"]),         difficultyPropId,  difficultyType?.id   ?? null);
      rel(splitVal(ex["Exercise type"]),      exTypePropId,      exTypeType?.id       ?? null);
      rel(splitVal(ex["Related topics"]),     relTopicPropId,    relatedTopicType?.id ?? null);
      rel(splitVal(ex["Training category"]),  trainingCatPropId, trainingCatType?.id  ?? null);
      rel(splitVal(ex["Variations"]),         variationsPropId,  variationType?.id    ?? null);
    }

    if ((i + 1) % 10 === 0) console.log(`  ${i + 1}/${batch.length} processed`);
  }

  console.log(`  New: ${newCount} | Patched (missing props filled): ${patchedCount}`);

  if (allOps.length === 0) {
    console.log("\n✓ Nothing new to upload — all entities already exist.");
    return;
  }

  // ── 7. PUBLISH ─────────────────────────────────────────────────────────────
  console.log(`\nTotal ops: ${allOps.length}`);
  console.log("Publishing to IPFS...");

  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: "Exercise Catalog — Batch 2 (exercises 51–100)",
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
