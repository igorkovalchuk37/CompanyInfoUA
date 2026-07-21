from __future__ import annotations

import argparse
import json
import math
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from lxml import etree

SVG_NS = 'http://www.w3.org/2000/svg'
Q = lambda tag: f'{{{SVG_NS}}}{tag}'
NS = {'svg': SVG_NS}


def project(lat: float, lon: float) -> tuple[float, float]:
    """Project WGS84 latitude/longitude to the template's 1000×669 coordinate space."""
    x = 50.44295494 * lon - 1071.19389230
    merc = math.log(math.tan(math.pi / 4 + math.radians(lat) / 2))
    y = -2890.33445159 * merc + 3142.28933775
    return x, y


def require_group(root: etree._Element, group_id: str) -> etree._Element:
    found = root.xpath(f"//*[local-name()='g' and @id='{group_id}']")
    if not found:
        raise ValueError(f'У шаблоні відсутній обов’язковий шар: {group_id}')
    return found[0]


def clear_group(group: etree._Element) -> None:
    for child in list(group):
        group.remove(child)


def add_text(parent: etree._Element, x: float, y: float, text: str, css_class: str, anchor: str = 'start') -> etree._Element:
    el = etree.SubElement(parent, Q('text'), {
        'x': f'{x:.1f}',
        'y': f'{y:.1f}',
        'class': css_class,
        'text-anchor': anchor,
    })
    el.text = text
    return el


def get_asset_xy(asset: dict[str, Any], manifest: dict[str, Any]) -> tuple[float, float, str]:
    if 'x' in asset and 'y' in asset:
        return float(asset['x']), float(asset['y']), 'explicit'
    if 'lat' in asset and 'lon' in asset:
        x, y = project(float(asset['lat']), float(asset['lon']))
        return x, y, 'latlon'
    rid = asset.get('region_id')
    if rid and rid in manifest['regions']:
        anchor = manifest['regions'][rid]['anchor']
        return float(anchor['x']), float(anchor['y']), 'region_anchor'
    raise ValueError(f"Актив '{asset.get('name', 'без назви')}' не має x/y, lat/lon або коректного region_id")


def collision_offset(index: int, total: int) -> tuple[float, float]:
    if total <= 1:
        return 0.0, 0.0
    radius = 12 + 5 * (index // 8)
    angle = 2 * math.pi * (index % 8) / min(total, 8) - math.pi / 2
    return radius * math.cos(angle), radius * math.sin(angle)


def render(data_path: Path, output_svg: Path, template_path: Path, manifest_path: Path, output_png: Path | None = None) -> None:
    data = json.loads(data_path.read_text(encoding='utf-8'))
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    tree = etree.parse(str(template_path))
    root = tree.getroot()

    if root.get('data-template') != 'ua-business-assets-map':
        raise ValueError('Переданий SVG не є підтримуваним шаблоном карти активів України')

    markers_g = require_group(root, 'asset-markers')
    labels_g = require_group(root, 'asset-labels')
    connections_g = require_group(root, 'asset-connections')
    title_g = require_group(root, 'map-title-layer')
    legend_g = require_group(root, 'map-legend')
    for g in (markers_g, labels_g, connections_g, title_g, legend_g):
        clear_group(g)

    # Reset and apply region highlights.
    for path in root.xpath("//*[local-name()='g' and @id='map-regions']/*[local-name()='path']"):
        path.set('class', 'region')

    assets = data.get('assets', [])
    active_regions = set(data.get('active_regions', []))
    if data.get('derive_active_regions', True):
        active_regions.update(a.get('region_id') for a in assets if a.get('region_id'))
    for rid in sorted(active_regions):
        found = root.xpath(f"//*[local-name()='path' and @id='{rid}']")
        if found:
            found[0].set('class', 'region is-active')

    # Resolve all positions first, then spread collisions.
    resolved = []
    bucket_counts: Counter[str] = Counter()
    for i, asset in enumerate(assets):
        x, y, source = get_asset_xy(asset, manifest)
        key = asset.get('region_id') if source == 'region_anchor' else f'{round(x, 1)}|{round(y, 1)}'
        bucket_counts[key] += 1
        resolved.append({'asset': asset, 'x0': x, 'y0': y, 'source': source, 'bucket': key, 'index': i})

    bucket_seen: defaultdict[str, int] = defaultdict(int)
    used_types = []
    for item in resolved:
        asset = item['asset']
        key = item['bucket']
        n = bucket_seen[key]
        bucket_seen[key] += 1
        dx, dy = collision_offset(n, bucket_counts[key])
        x = item['x0'] + dx
        y = item['y0'] + dy

        asset_type = asset.get('type', 'other')
        if asset_type not in manifest['asset_types']:
            asset_type = 'other'
        if asset_type not in used_types:
            used_types.append(asset_type)

        marker = etree.SubElement(markers_g, Q('g'), {
            'class': f'asset-marker asset-{asset_type}',
            'transform': f'translate({x:.2f} {y:.2f})',
            'data-asset-id': str(asset.get('id', item['index'] + 1)),
            'data-asset-type': asset_type,
        })
        use = etree.SubElement(marker, Q('use'), {'href': '#asset-pin'})
        use.set('aria-hidden', 'true')
        title = etree.SubElement(marker, Q('title'))
        title.text = asset.get('name', 'Актив')

        if abs(dx) > 0.1 or abs(dy) > 0.1:
            etree.SubElement(connections_g, Q('line'), {
                'class': 'connector',
                'x1': f"{item['x0']:.1f}", 'y1': f"{item['y0']:.1f}",
                'x2': f'{x:.1f}', 'y2': f'{y:.1f}',
            })

        if asset.get('show_label', True):
            label = asset.get('name', 'Актив')
            side = asset.get('label_side', 'auto')
            estimated_width = max(45.0, len(str(label)) * 7.4)
            if side == 'auto':
                side = 'left' if x + 13 + estimated_width > 990 else 'right'
            if side == 'left' and x - 13 - estimated_width < 10:
                side = 'right'
            if side == 'right' and x + 13 + estimated_width > 990:
                side = 'left'
            if side == 'left':
                tx, anchor = x - 13, 'end'
            else:
                tx, anchor = x + 13, 'start'
            ty = y - 14 + float(asset.get('label_dy', 0))
            add_text(labels_g, tx, ty, label, 'asset-label-halo', anchor)
            add_text(labels_g, tx, ty, label, 'asset-label', anchor)
            sub = asset.get('city') or asset.get('description')
            if sub and asset.get('show_sublabel', True):
                add_text(labels_g, tx, ty + 15, str(sub), 'asset-sublabel-halo', anchor)
                add_text(labels_g, tx, ty + 15, str(sub), 'asset-sublabel', anchor)

    if data.get('show_title', False):
        group_name = data.get('group_name', 'Карта активів бізнес-групи')
        add_text(title_g, 28, 36, group_name, 'map-title')
        subtitle = data.get('subtitle')
        if subtitle:
            add_text(title_g, 28, 56, subtitle, 'map-subtitle')

    if data.get('show_legend', False) and used_types:
        x0, y0 = 28.0, 500.0
        row_h = 24.0
        width = 210.0
        height = 18.0 + row_h * len(used_types)
        etree.SubElement(legend_g, Q('rect'), {
            'class': 'legend-bg', 'x': f'{x0:.1f}', 'y': f'{y0:.1f}',
            'width': f'{width:.1f}', 'height': f'{height:.1f}', 'rx': '8'
        })
        for idx, asset_type in enumerate(used_types):
            y = y0 + 24 + idx * row_h
            mark = etree.SubElement(legend_g, Q('g'), {
                'class': f'asset-marker asset-{asset_type}',
                'transform': f'translate({x0 + 18:.1f} {y:.1f}) scale(.7)',
            })
            etree.SubElement(mark, Q('use'), {'href': '#asset-dot'})
            add_text(legend_g, x0 + 34, y + 4, manifest['asset_types'][asset_type]['label_uk'], 'legend-label')

    # Update accessible title/description.
    title_el = root.xpath("//*[local-name()='title' and @id='map-title-element']")[0]
    title_el.text = data.get('group_name') or 'Карта активів бізнес-групи в Україні'
    desc_el = root.xpath("//*[local-name()='desc' and @id='map-description-element']")[0]
    desc_el.text = data.get('description') or f"Карта містить {len(assets)} активів бізнес-групи."

    output_svg.parent.mkdir(parents=True, exist_ok=True)
    tree.write(str(output_svg), xml_declaration=True, encoding='UTF-8', pretty_print=True)

    if output_png:
        try:
            import cairosvg
        except ImportError as exc:
            raise RuntimeError('Для PNG встановіть cairosvg: pip install cairosvg') from exc
        cairosvg.svg2png(url=str(output_svg), write_to=str(output_png), output_width=1600)


def main() -> None:
    parser = argparse.ArgumentParser(description='Створення карти активів бізнес-групи на основі стабільного SVG-шаблону України.')
    parser.add_argument('data_json', type=Path)
    parser.add_argument('output_svg', type=Path)
    parser.add_argument('--template', type=Path, default=Path(__file__).with_name('ua_business_map_template.svg'))
    parser.add_argument('--manifest', type=Path, default=Path(__file__).with_name('ua_business_map_manifest.json'))
    parser.add_argument('--png', type=Path, default=None)
    args = parser.parse_args()
    render(args.data_json, args.output_svg, args.template, args.manifest, args.png)


if __name__ == '__main__':
    main()
