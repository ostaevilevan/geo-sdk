import "dotenv/config";
import {
  Graph,
  personalSpace,
  getSmartAccountWalletClient,
  SystemIds,
} from "@geoprotocol/geo-sdk";

// Your personal GEO space ID
const SPACE_ID = "0c747ad3af58ed6b27221a256498068e";

async function main() {
  const privateKey = process.env.PRIVATE_KEY as `0x${string}`;
  if (!privateKey || privateKey.includes("paste_your")) {
    throw new Error("Set your PRIVATE_KEY in the .env file first.");
  }

  // 1. Create a simple test entity using built-in SDK system IDs
  const entity = Graph.createEntity({
    name: "SDK Test Entity",
    types: [],
    values: [
      {
        property: SystemIds.NAME_PROPERTY,
        type: "text",
        value: "SDK Test Entity",
      },
      {
        property: SystemIds.DESCRIPTION_PROPERTY,
        type: "text",
        value: "A test entity uploaded via the GEO SDK to verify the setup works correctly.",
      },
    ],
  });

  const allOps = [...entity.ops];
  console.log(`Entity ID: ${entity.id}`);
  console.log(`Total ops: ${allOps.length}`);

  // 2. Publish ops to IPFS and get on-chain calldata
  console.log("Publishing edit to IPFS...");
  const { editId, to, calldata } = await personalSpace.publishEdit({
    name: "Test Upload — SDK setup check",
    spaceId: SPACE_ID,
    ops: allOps,
    author: SPACE_ID, // your personal space is also the author
    network: "TESTNET",
  });
  console.log(`Edit ID: ${editId}`);

  // 3. Get smart account wallet client (gas-sponsored, no ETH needed)
  console.log("Getting smart account wallet...");
  const walletClient = await getSmartAccountWalletClient({ privateKey });

  // 4. Send the transaction on-chain
  console.log("Sending transaction...");
  const txHash = await walletClient.sendTransaction({ to, data: calldata });

  console.log(`✓ Success!`);
  console.log(`  Edit ID : ${editId}`);
  console.log(`  Tx Hash : ${txHash}`);
  console.log(`  View    : https://geobrowser.io/space/${SPACE_ID}`);
}

main().catch((err) => {
  console.error("✗ Error:", err.message ?? err);
  process.exit(1);
});
