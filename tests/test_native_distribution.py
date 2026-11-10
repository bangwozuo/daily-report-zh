"""Native distribution and source compatibility; no provider or business call."""
from pathlib import Path
import hashlib,json
import yaml

ROOT=Path(__file__).resolve().parents[1]
PACKAGE=json.loads((ROOT/'package.json').read_text('utf-8'))
EMPLOYEE=PACKAGE['employees'][0]

def test_original_package_attributes_retained():
    package=yaml.safe_load((ROOT/'package.yaml').read_text('utf-8'))
    assert set(package['legacy_package_definition'])<=set(package)
    assert package['legacy_name']=='门店经营日报分析师'
    assert package['demand_ids']==['D02','D06','D27']

def test_legacy_method_interfaces_unchanged():
    binding=json.loads((ROOT/'SOURCE-BINDING.json').read_text('utf-8'))
    for row in binding['compatibility_repair']['restored_interfaces']:
        assert hashlib.sha256((ROOT/row['path']).read_bytes()).hexdigest()==row['original_sha256']
        assert (ROOT/row['path']).with_name(row['native_file']).is_file()

def test_native_calls_target_exact_existing_assets():
    assets=json.loads((ROOT/'server/bindings/original-seven-assets.json').read_text('utf-8'))
    assert len(assets)==7
    targets={a['action']:a for a in assets}
    calls=[a for a in EMPLOYEE['skills']if a['kind']=='workflow_call']
    assert {a['execution']['operation']for a in calls}=={'collect','daily','alert','weekly'}
    for call in calls:
        flow=json.loads((ROOT/call['distribution']['path']/'workflow.json').read_text('utf-8'))
        assert flow['native_release']==targets[flow['action']]
        assert flow['host_creates_second_checkpoint']is False
        assert call['execution']['checkpoint_authority']=='existing WorkflowService'

def test_main_and_workflow_clients_share_one_transport():
    sdk=(ROOT/'clients/sdk/daily.py').read_bytes()
    assert (ROOT/'clients/skill/daily-report-zh/scripts/product.py').read_bytes()==sdk
    for row in EMPLOYEE['skills']:
        assert (ROOT/row['distribution']['path']/'scripts/product.py').read_bytes()==sdk
        assert row['execution']['service_required']is True
    assert EMPLOYEE['distribution']['categories']=={'employee':1,'generic':3,'workflow':4,'remote_service':0}

def test_eight_current_knowledge_bodies_and_limits():
    sources=json.loads((ROOT/'server/bindings/current-knowledge-sources.json').read_text('utf-8'))
    assert len(sources)==8
    for source in sources:
        assert hashlib.sha256((ROOT/source['source_file']).read_bytes()).hexdigest()==source['source_sha256']
        assert source['source_revision'] and source['generation_id']
    assert EMPLOYEE['qualification']['full_191_release_accepted']is False
    assert EMPLOYEE['qualification']['R78']=='retained'
    assert PACKAGE['quality_report']['full_191_release_accepted']is False
