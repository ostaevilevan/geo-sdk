import "dotenv/config";
import XLSX from "xlsx";
import {
  Graph,
  personalSpace,
  getSmartAccountWalletClient,
  SystemIds,
} from "@geoprotocol/geo-sdk";

const SPACE_ID = "0c747ad3af58ed6b27221a256498068e";
const EXCEL_PATH = "/Users/levan/Desktop/My best friend/Exercises Catalog (1).xlsx";
const BATCH_LIMIT = 50;

function split(val: unknown): string[] {
  if (!val || typeof val !== "string") return [];
  return val.split(",").map((s) => s.trim()).filter(Boolean);
}

async function main() {
  const privateKey = process.env.PRIVATE_KEY as `0x${string}`;
  if (!privateKey?.startsWith("0x")) throw new Error("Invalid PRIVATE_KEY in .env");

  const allOps: any[] = [];

  // ── SCHEMA ──────────────────────────────────────────────────────────────────
  console.log("Creating schema...");

  // Custom types
  const exerciseType        = Graph.createType({ name: "Exercise" });
  const bodySystemType      = Graph.createType({ name: "Body system" });
  const difficultyType      = Graph.createType({ name: "Difficulty level" });
  const exerciseTypeEntity  = Graph.createType({ name: "Exercise type" });
  const trainingCatType     = Graph.createType({ name: "Training category" });
  const relatedTopicType    = Graph.createType({ name: "Related topic" });
  const variationType       = Graph.createType({ name: "Exercise variation" });

  // Custom relation properties
  const bodySystemProp      = Graph.createProperty({ name: "Body system",       dataType: "RELATION" });
  const difficultyProp      = Graph.createProperty({ name: "Difficulty",         dataType: "RELATION" });
  const exerciseTypeProp    = Graph.createProperty({ name: "Exercise type",      dataType: "RELATION" });
  const trainingCatProp     = Graph.createProperty({ name: "Training category",  dataType: "RELATION" });
  const variationsProp      = Graph.createProperty({ name: "Variations",         dataType: "RELATION" });
  const relatedTopicProp    = Graph.createProperty({ name: "Related topics",     dataType: "RELATION" });

  // Custom text properties
  const commonMistakesProp  = Graph.createProperty({ name: "Common mistakes",   dataType: "TEXT" });
  const equipmentProp       = Graph.createProperty({ name: "Equipment",          dataType: "TEXT" });
  const formCuesProp        = Graph.createProperty({ name: "Form cues",          dataType: "TEXT" });
  const primaryMusclesProp  = Graph.createProperty({ name: "Primary muscles",    dataType: "TEXT" });
  const secondaryMusclesProp = Graph.createProperty({ name: "Secondary muscles", dataType: "TEXT" });

  for (const r of [
    exerciseType, bodySystemType, difficultyType, exerciseTypeEntity,
    trainingCatType, relatedTopicType, variationType,
    bodySystemProp, difficultyProp, exerciseTypeProp, trainingCatProp,
    variationsProp, relatedTopicProp,
    commonMistakesProp, equipmentProp, formCuesProp,
    primaryMusclesProp, secondaryMusclesProp,
  ]) {
    allOps.push(...r.ops);
  }

  // ── READ EXCEL ───────────────────────────────────────────────────────────────
  console.log("Reading Excel...");
  const wb = XLSX.readFile(EXCEL_PATH);

  const exercises         = XLSX.utils.sheet_to_json<any>(wb.Sheets["Exercises"]);
  const bodySystemRows    = XLSX.utils.sheet_to_json<any>(wb.Sheets["Body System"]);
  const difficultyRows    = XLSX.utils.sheet_to_json<any>(wb.Sheets["Difficulty"]);
  const exerciseTypeRows  = XLSX.utils.sheet_to_json<any>(wb.Sheets["Exercise type"]);
  const relatedTopicRows  = XLSX.utils.sheet_to_json<any>(wb.Sheets["Related topics"]);
  const trainingCatRows   = XLSX.utils.sheet_to_json<any>(wb.Sheets["Training Category - minus other"]);
  const variationRows     = XLSX.utils.sheet_to_json<any>(wb.Sheets["Variations"]);

  const batch = exercises.slice(0, BATCH_LIMIT);

  // Collect variation names referenced by this batch
  const referencedVariations = new Set<string>();
  batch.forEach((ex) => split(ex["Variations"]).forEach((v) => referencedVariations.add(v)));

  // ── SUPPORTING ENTITIES ──────────────────────────────────────────────────────
  console.log("Creating supporting entities...");

  const idMap = new Map<string, string>(); // entity name → generated ID

  function createSupporting(name: string, description: string, typeId: string) {
    if (!name || idMap.has(name)) return;
    const e = Graph.createEntity({
      name,
      types: [typeId],
      values: [
        { property: SystemIds.NAME_PROPERTY,        type: "text", value: name },
        { property: SystemIds.DESCRIPTION_PROPERTY, type: "text", value: description },
      ].filter((v) => v.value),
    });
    idMap.set(name, e.id);
    allOps.push(...e.ops);
  }

  bodySystemRows.forEach((r) =>
    createSupporting(r["Entity Name"], r["Description"] ?? "", bodySystemType.id));

  difficultyRows.forEach((r) =>
    createSupporting(r["Entity Name"], r["Description"] ?? "", difficultyType.id));

  exerciseTypeRows.forEach((r) =>
    createSupporting(r["Entity Name"], r["Description"] ?? "", exerciseTypeEntity.id));

  relatedTopicRows.forEach((r) =>
    createSupporting(r["Entity Name"], r["Description"] ?? "", relatedTopicType.id));

  trainingCatRows
    .filter((r) => r["Name"])
    .forEach((r) =>
      createSupporting(r["Name"], r["Clear definition"] ?? "", trainingCatType.id));

  variationRows
    .filter((r) => r["Entity Name"] && referencedVariations.has(r["Entity Name"]))
    .forEach((r) =>
      createSupporting(r["Entity Name"], r["Description"] ?? "", variationType.id));

  console.log(`Supporting entities created: ${idMap.size}`);

  // ── EXERCISES + RELATIONS ────────────────────────────────────────────────────
  console.log(`Creating ${batch.length} exercises...`);

  batch.forEach((ex, i) => {
    const name = ex["Entity Name"] as string;

    const exercise = Graph.createEntity({
      name,
      types: [exerciseType.id],
      values: [
        { property: SystemIds.NAME_PROPERTY,        type: "text", value: name },
        { property: SystemIds.DESCRIPTION_PROPERTY, type: "text", value: ex["Description"] ?? "" },
        { property: commonMistakesProp.id,  type: "text", value: ex["Common mistakes"] ?? "" },
        { property: equipmentProp.id,       type: "text", value: ex["Equipment"] ?? "" },
        { property: formCuesProp.id,        type: "text", value: ex["Form cues"] ?? "" },
        { property: primaryMusclesProp.id,  type: "text", value: ex["Primary muscles"] ?? "" },
        { property: secondaryMusclesProp.id,type: "text", value: ex["Secondary muscles"] ?? "" },
      ].filter((v) => v.value),
    });

    allOps.push(...exercise.ops);
    idMap.set(name, exercise.id);

    // Create relation helper
    const rel = (toName: string, propId: string) => {
      const toId = idMap.get(toName);
      if (!toId) { console.warn(`  ⚠ No entity found for: "${toName}"`); return; }
      const r = Graph.createRelation({ fromEntity: exercise.id, toEntity: toId, type: propId });
      allOps.push(...r.ops);
    };

    split(ex["Body systems"]).forEach((v)      => rel(v, bodySystemProp.id));
    split(ex["Difficulty"]).forEach((v)         => rel(v, difficultyProp.id));
    split(ex["Exercise type"]).forEach((v)      => rel(v, exerciseTypeProp.id));
    split(ex["Related topics"]).forEach((v)     => rel(v, relatedTopicProp.id));
    split(ex["Training category"]).forEach((v)  => rel(v, trainingCatProp.id));
    split(ex["Variations"]).forEach((v)         => rel(v, variationsProp.id));

    if ((i + 1) % 10 === 0) console.log(`  ${i + 1}/${batch.length} done`);
  });

  // ── PUBLISH ──────────────────────────────────────────────────────────────────
  console.log(`\nTotal ops: ${allOps.length}`);
  console.log("Publishing to IPFS...");

  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: "Exercise Catalog — Batch 1 (50 exercises)",
    spaceId: SPACE_ID,
    ops: allOps,
    author: SPACE_ID,
    network: "TESTNET",
  });

  console.log("Getting wallet...");
  const walletClient = await getSmartAccountWalletClient({ privateKey });

  console.log("Sending transaction...");
  const txHash = await walletClient.sendTransaction({ to, data: calldata });

  console.log(`\n✓ Success!`);
  console.log(`  Edit ID  : ${editId}`);
  console.log(`  Tx Hash  : ${txHash}`);
  console.log(`  View     : https://geobrowser.io/space/${SPACE_ID}`);
}

main().catch((err) => {
  console.error("✗ Error:", err.message ?? err);
  process.exit(1);
});
