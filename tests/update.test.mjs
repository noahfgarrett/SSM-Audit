import test from 'node:test'
import assert from 'node:assert/strict'
import { compareVersions, selectReleaseAsset, versionedFilename } from '../src/update/update.js'

test('version comparison handles semantic release numbers',()=>{
  assert.equal(compareVersions('1.0.1','1.0.0'),1)
  assert.equal(compareVersions('1.0.0','1.0.0'),0)
  assert.equal(compareVersions('1.2.0','2.0.0'),-1)
})

test('updater selects only the standalone versioned HTML asset',()=>{
  const info=selectReleaseAsset({tag_name:'v1.2.3',body:'notes',assets:[
    {name:'SSM-Audit-v1.2.3.html.gz',browser_download_url:'gzip'},
    {name:'SSManagement-v4.3.0.html',browser_download_url:'wrong'},
    {name:'SSM-Audit-v1.2.3.html',browser_download_url:'right'},
  ]})
  assert.equal(info.downloadUrl,'right')
  assert.equal(info.version,'1.2.3')
  assert.equal(versionedFilename(info.version),'SSM-Audit-v1.2.3.html')
})

test('updater rejects an HTML asset whose version does not match the release tag',()=>{
  assert.equal(selectReleaseAsset({tag_name:'v1.2.3',assets:[{name:'SSM-Audit-v1.2.2.html',browser_download_url:'wrong'}]}),null)
})

test('updater accepts a latest.json manifest only when the download URL names the versioned asset',async()=>{
  const { selectLatestManifest, latestManifestUrl } = await import('../src/update/update.js')
  const good=selectLatestManifest({version:'v1.12.0',downloadUrl:'https://example.test/releases/download/v1.12.0/SSM-Audit-v1.12.0.html',releaseNotes:'- x'})
  assert.equal(good.version,'1.12.0');assert.equal(good.assetName,'SSM-Audit-v1.12.0.html')
  assert.equal(selectLatestManifest({version:'1.12.0',downloadUrl:'https://example.test/SSM-Audit-v1.11.0.html'}),null,'asset must match the manifest version')
  assert.equal(selectLatestManifest({version:'latest',downloadUrl:'x'}),null)
  assert.equal(latestManifestUrl('owner/repo'),'https://raw.githubusercontent.com/owner/repo/main/latest.json')
})

test('updater tells a rate-limited check apart from an unreachable one',async()=>{
  const { classifyCheckFailure, describeUpdateCheck } = await import('../src/update/update.js')
  assert.equal(classifyCheckFailure(403,new Headers({'x-ratelimit-remaining':'0'})),'rate-limited')
  assert.equal(classifyCheckFailure(429,new Headers()),'rate-limited')
  assert.equal(classifyCheckFailure(500,new Headers()),'unreachable')
  assert.equal(classifyCheckFailure(0,null),'unreachable')
  assert.match(describeUpdateCheck({state:'rate-limited'},'1.0.0'),/limiting update checks/)
  assert.match(describeUpdateCheck({state:'current'},'1.0.0'),/latest version \(v1\.0\.0\)/)
  assert.match(describeUpdateCheck({state:'update',version:'1.2.0'},'1.0.0'),/1\.2\.0 is available/)
  assert.equal(describeUpdateCheck(null,'1.0.0'),'Not checked yet.')
})
