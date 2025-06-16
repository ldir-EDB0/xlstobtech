#!/usr/bin/env node

// inspired by script written by Bridgetech that did one csv file
// Extensively modified, mangled and extended by Kevin Darbyshire-Bryant
// with significant aid from ChatGPT & Copilot

const xm = require('xmlbuilder2');

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Parse command line arguments
function parseArgs() {
  const args = process.argv.slice(2);
  let inputFile;
  let outputDir = 'btechxml';
  let pushProbe = '';
  let pushSheet = '';
  let pushMode = '';

  for (let i = 0; i < args.length; i++) {
    if (args[i] === '-x' && i + 1 < args.length) {
      inputFile = args[i + 1];
      i++;
    } else if (args[i] === '-d' && i + 1 < args.length) {
      outputDir = args[i + 1];
      i++;
    } else if (args[i] === '-p' && i + 1 < args.length) {
      pushProbe = args[i + 1];
      i++;
    } else if (args[i] === '-s' && i + 1 < args.length) {
      pushSheet = args[i + 1];
      i++;
    } else if (args[i] === '-u') {
      pushMode = 'update';
    } else if (args[i] === '-r') {
      pushMode = 'delete';
    }
  }
  if (!inputFile || (Boolean(pushProbe) !== Boolean(pushSheet))) {
    console.log(`
Usage:
  node xlstobtech.js -x <input.xls> [ [-d <output-folder>] || -p <probe> -s <sheet> ]

Options:
  -x <file>      Excel file to process (required)

  -d <folder>    Output directory (optional, defaults to btechxml)
or
  -s <sheet>     Specify sheet to push
  -p <probe>     Push specified sheet to probe

  Examples:
  node xlstobtech.js -x input.xlsx
  node xlstobtech.js -x input.xls -d ./configs
`);
    process.exit(0);
  }

  console.log(`Input XLS: ${inputFile}`);
  console.log(`Output directory: ${outputDir}`);
  
  return { inputFile, outputDir, pushProbe, pushSheet, pushMode };
}

// Process unicast sheet and extract interface data
function processUnicastSheet(workbook) {
  const sheet = workbook.Sheets['unicast'];
  if (!sheet) throw new Error('⚠️  Error: "unicast" sheet not found');

  let json = xlsx.utils.sheet_to_json(sheet, { defval: '' });
  if (json.length === 0) throw new Error('⚠️  Error: "unicast" sheet is empty');

  const interfaceNames = [];
  const interfaceByNameVlan = {};

  for (const row of json) {
    const Name = row['FRIENDLY_NAME'];
    if (!Name) continue;

    const Vlan = row['VLAN'] ? row['VLAN'].toString().toLowerCase() : '';
    const Interface = row['INTERFACE'];
    const IPGW = row['IP_PRFX_GW'];

    if (Name && !interfaceNames.includes(Name)) interfaceNames.push(Name);

    if (Name && Interface) {
      const key = `${Name}-${Vlan}`;
      if (!interfaceByNameVlan[key]) {
        interfaceByNameVlan[key] = {
          Interface,
          IPGW
        };
      }
    }
  }

  console.log(`Found ${interfaceNames.length} Probes in unicast sheet`);
  return { interfaceNames, interfaceByNameVlan };
}

// Process profiles sheet
function processProfilesSheet(workbook) {
  const sheet = workbook.Sheets['profiles'];
  if (!sheet) throw new Error('⚠️  Error: "profiles" sheet not found, I need some data from it.');
  
  const json = xlsx.utils.sheet_to_json(sheet, { defval: '' });
  if (json.length === 0) throw new Error('⚠️  Error: "profiles" sheet is empty, I need some data from it.');

  const profiles = {};
  
  for (const row of json) {
    const profile = row['profile'];
    if (profile) {
      profiles[profile] = {
        profile,
        content: row['content'] || '',
        audiodepth: row['audiodepth'] || '24',
        channelorder: row['channelorder'] || 'ST',
        audiosr: row['audiosr'] || '48000',
        port_no_a: row['port_no_a'] || '5004',
        port_no_b: row['port_no_b'] || '5004',
        notes: row['Notes'] || ''
      };
    }
  }
  
  console.log(`Found ${Object.keys(profiles).length} profiles in profiles sheet`);
  return profiles;
}

function buildMcastChannel(name, source_ip, multicast, port, iface, profile, groups, page, join) {
  return {
    name: name,
    addr: multicast,
    port: port,
    sessionId: "0",
    groups: groups,
    audiodepth: profile.audiodepth,
    audiosr: profile.audiosr,
    channelOrder: profile.channelorder,
    joinIfaceName: iface,
    ssmAddr: source_ip,
    join: join,
    page: page,
    etrEngine: "1",
    extractThumbs: true,
    enableFec: false,
    enableRtcp: true
  };
}

// Process individual sheet and generate multicasts
function processSheet(workbook, sheetName, probe, interfaceByNameVlan, profiles) {
  console.log(`🔄 Processing sheet: ${sheetName}`);
  const sheet = workbook.Sheets[sheetName];
  const json = xlsx.utils.sheet_to_json(sheet, { defval: '' });

  if (json.length === 0) {
    console.warn(`⚠️  Skipping empty sheet: ${sheetName}`);
    return null;
  }

  const multicasts = [];
  let skipped = 0;

  for (const row of json) {
    const groups = row['groups'] || '';
    const page = row['page'] || '1';
    const name = row['name'] || '';
    const device = row['device'] || '';
    const join = (row['join'] && row['join'].toString().trim().toLowerCase() === 'no') ? false : true;
    const profileName = row['profile'] || '';
    const source_ip_a = row['source_ip_a'] || '';
    const multicast_a = row['multicast_a'] || '';
    const vlan_a = row['vlan_a'] || 'dff-a';
    const source_ip_b = row['source_ip_b'] || '';
    const multicast_b = row['multicast_b'] || '';
    const vlan_b = row['vlan_b'] || 'dff-b';

    // Lookup profile data
    const profile = profiles[profileName] || {};
  
    // Get probe interface name for VLAN
    const iface_a = interfaceByNameVlan[`${probe}-${vlan_a}`];
    const iface_b = interfaceByNameVlan[`${probe}-${vlan_b}`];

    //A leg
    if (iface_a && source_ip_a && multicast_a) {
      const mname = source_ip_b ? `${name}@A` : `${name}`;

      multicasts.push(buildMcastChannel(mname, source_ip_a, multicast_a, profile.port_no_a, iface_a.Interface, profile, groups, page, join));
    }

    //B leg
    if (iface_b && source_ip_b && multicast_b) {
      const mname = source_ip_a ? `${name}@B` : `${name}`;

      multicasts.push(buildMcastChannel(mname, source_ip_b, multicast_b, profile.port_no_b, iface_b.Interface, profile, groups, page, join));
    }

    if ((!multicast_a && !multicast_b) || (!source_ip_a && !source_ip_b)) {
      skipped++;
      continue;
    }
  }

  console.log(`✅ Sheet "${sheetName}": ${multicasts.length} entries (skipped ${skipped})`);
  return multicasts.length ? multicasts : null;
}

// write XML config for BTech probe
function wrapXml(multicasts) {
  const doc = xm.create({ version: '1.0' })
    .ele('ewe', {
      mask: '0x80',
      hw_type: '440',
      br: 'BT'
    })
      .ele('probe')
        .ele('core')
          .ele('setup')
            .ele('mcastnames')
              .ele('mclist', { xmlChildren: 'list' });

  multicasts.forEach(mc =>
    doc.ele('mcastChannel', mc)
  );

  return doc.end({ prettyPrint: true });
}

// Write output xml file
function writeConfigFile(outputDir, probe, sheetName, multicasts) {

  const btechxml = wrapXml(multicasts);
  const safeSheetName = sheetName.replace(/[ \\/:*?"<>|]/g, '_');
  const safeprobe = probe.replace(/[ \\/:*?"<>|]/g, '_');
  const probeoutputDir = path.join(outputDir, safeprobe);
  fs.mkdirSync(probeoutputDir, { recursive: true });
  const outputPath = path.join(probeoutputDir, `${safeSheetName}.xml`);

  fs.writeFileSync(outputPath, btechxml);
  console.log(`💾 Written: ${outputPath}`);
}

// Push config to specified probe
async function pushConfig(interfaceByNameVlan, probe, sheetName, pushMode, multicasts) {

  const iface = interfaceByNameVlan[`${probe}-dtv`];
  if (!iface) {
    console.error(`❌ Error: DTV Interface for probe "${probe}" not found.`);
    return;
  }

  // Extract the IP address from the IPGW string (format: ip/prefix/gateway)
  const probeIpAddress = iface.IPGW.split('/')[0];
  if (!probeIpAddress) {
    console.error(`❌ Error: IP address for probe "${probe}" not found.`);
    return;
  }

  const btechxml = wrapXml(multicasts);

  const probeUrl = new URL(pushMode ? `http://${probeIpAddress}/probe/core/importExport/data.xml?mode=${pushMode}` : `http://${probeIpAddress}/probe/core/importExport/data.xml`);

  console.log(`📤 Pushing config for ${sheetName} to probe ${probe} at ${probeIpAddress} using URL ${probeUrl}`);

  try {
    const res = await fetch(probeUrl, {
       method: 'POST',
       headers: { 'Content-Type': 'application/xml' },
       body: btechxml,
    });

    if (!res.ok) {
      const errorText = await res.text();
      throw new Error(`HTTP ${res.status}: ${errorText}`);
    }

    console.log(`✅ Successfully uploaded XML to ${probe}`);
  } catch (postErr) {
    console.error(`❌ Failed to POST to ${probe}:`, postErr.message);
  }
}

function ProcessAllSheets(workbook, sheetNames, Probes, interfaceByNameVlan, profiles, outputDir) {
  // Produce a config file for each probe from each sheet

  //create base output directory if it doesn't exist
  fs.mkdirSync(outputDir, { recursive: true });

  for (const sheetName of sheetNames) {
    for (const probe of Probes) {
      const multicasts = processSheet(workbook, sheetName, probe, interfaceByNameVlan, profiles);
      if (multicasts) writeConfigFile(outputDir, probe, sheetName, multicasts);
    }
  }
}

// Main function
async function main() {
  const skipSheets = new Set(['unicast', 'profiles', 'validation']);

  try {
    const { inputFile, outputDir, pushProbe, pushSheet, pushMode } = parseArgs();

    console.log(`Reading Excel file...`);
    const workbook = xlsx.readFile(inputFile);
    // Exclude sheets in skipSheets from sheetNames
    const sheetNames = workbook.SheetNames.filter(name => !skipSheets.has(name));
    if (sheetNames.length === 0) throw new Error('No valid sheets to process in Excel file.');

    //get the probe names and their interfaces from the unicast sheet
    const { interfaceNames: Probes, interfaceByNameVlan } = processUnicastSheet(workbook);

    //get audio/video profiles from the profiles sheet
    const profiles = processProfilesSheet(workbook);

    // If pushProbe and pushSheet are specified, process only that sheet for the probe
    if (pushProbe && pushSheet) {
      if (!Probes.includes(pushProbe)) {
        console.error(`❌ Error: Probe "${pushProbe}" not found in unicast sheet.`);
      }
      if (skipSheets.has(pushSheet)) {
        console.error(`❌ Error: Sheet "${pushSheet}" is not a pushable sheet.`);
      }
      else if (!sheetNames.includes(pushSheet)) {
        console.error(`❌ Error: Sheet "${pushSheet}" not found in the workbook.`);
      }
      
      const multicasts = processSheet(workbook, pushSheet, pushProbe, interfaceByNameVlan, profiles);
      if (multicasts) await pushConfig(interfaceByNameVlan, pushProbe, pushSheet, pushMode, multicasts);
      console.log(`🎉 Processed ${pushSheet} for probe ${pushProbe}.`);
    }
    else {
      ProcessAllSheets(workbook, sheetNames, Probes, interfaceByNameVlan, profiles, outputDir);
      console.log(`🎉 All sheets processed.`);
    }
  } catch (err) {
    console.error('❌ Error:', err.message);
    process.exit(1);
  }
}

// Execute the program
main();
