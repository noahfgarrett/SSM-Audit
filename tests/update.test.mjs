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
