const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const html = fs.readFileSync(path.join(root, 'route_map.html'), 'utf8');
const context = vm.createContext({ console, Number, Math, Map });

function loadConst(name) {
  const start = html.indexOf(`const ${name} = `);
  assert.notEqual(start, -1, `${name} declaration not found`);
  const tail = html.slice(start);
  const terminator = tail.search(/;\r?\nconst /);
  const end = terminator === -1 ? -1 : start + terminator;
  assert.notEqual(end, -1, `${name} declaration terminator not found`);
  vm.runInContext(html.slice(start, end + 1), context);
}

loadConst('TIMETABLE');
loadConst('FARE_PROFILES');

const routingStart = html.indexOf('let stationDepartures = {};');
const routingEnd = html.indexOf('let physicalAdj = {};', routingStart);
assert.notEqual(routingStart, -1, 'routing code start not found');
assert.notEqual(routingEnd, -1, 'routing code end not found');
vm.runInContext(html.slice(routingStart, routingEnd), context);
vm.runInContext('buildRoutingGraph();', context);

const result = vm.runInContext(
  `findFixedRoute(['부전', '청량리'], '01:53', 'FARE')`,
  context
);
assert.equal(result.error, undefined, result.error);
assert.ok(result.path.length > 0, 'route must contain at least one train');

const startTime = 1 * 60 + 53;
const totalDuration = result.finalTime - startTime;
assert.ok(totalDuration <= 24 * 60, `duration is still impractical: ${totalDuration} minutes`);
assert.ok(result.totalTransfers <= 2, `too many transfers: ${result.totalTransfers}`);

for (let index = 1; index < result.path.length; index += 1) {
  const wait = result.path[index].b_t - result.path[index - 1].a_t;
  assert.ok(wait <= 180, `transfer wait exceeds 180 minutes: ${wait}`);
}

console.log(JSON.stringify({
  arrival: result.finalTime,
  duration: totalDuration,
  transfers: result.totalTransfers,
  fare: result.totalFare,
  legs: result.path.map(leg => ({
    train: `${leg.trip_type} #${leg.trip_no}`,
    from: leg.b_st,
    departure: leg.b_t,
    to: leg.a_st,
    arrival: leg.a_t,
    fare: leg.fare
  }))
}, null, 2));
