"""
Kiwere Batch Map Export Script
================================
Run this in the QGIS Python Console (Plugins > Python Console).
"""

from qgis.core import (
    QgsProject, QgsLayoutExporter,
    QgsPrintLayout, QgsReadWriteContext,
    QgsLayerTreeLayer, QgsLegendStyle,
    QgsMapLayerLegendUtils, QgsLegendRenderer
)
from qgis.PyQt.QtXml import QDomDocument
from qgis.PyQt.QtCore import QFile, QIODevice
import os

# ========== CONFIGURATION ==========
QPT_PATH = r"C:/Users/user/Documents/Kiwere_Layout.qpt"
OUTPUT_DIR = r"C:/Users/user/Documents"

RANK_YEAR_MAP = {
    2: 2018,
    3: 2019,
    4: 2020,
    5: 2021,
    6: 2022,
    7: 2023,
    8: 2024,
    9: 2025,
}

EXPORT_DPI = 300
# ====================================


def find_raster_layer_for_rank(rank):
    project = QgsProject.instance()
    for layer in project.mapLayers().values():
        if f"Rank_{rank}_Year_" in layer.name() or f"Rank_{rank}_Year_" in layer.source():
            return layer
    return None


def load_qpt_template(qpt_path):
    f = QFile(qpt_path)
    if not f.open(QIODevice.ReadOnly):
        raise Exception(f"Cannot open QPT file: {qpt_path}")
    content = f.readAll().data().decode('utf-8')
    f.close()
    return content


def export_map_for_rank(rank, year, qpt_content, manager, project):
    target_layer = find_raster_layer_for_rank(rank)
    if target_layer is None:
        print(f"  WARNING: No layer found for Rank {rank} (Year {year}). Skipping.")
        return False

    print(f"  Found layer: {target_layer.name()}")

    modified_xml = qpt_content

    # ---- Update the HTML title from 2017 to the correct year ----
    modified_xml = modified_xml.replace(
        '&#xa;2017&#xa;',
        f'&#xa;{year}&#xa;'
    )

    # ---- Update legend layer references ----
    old_source = '/vsizip/../Downloads/drive-download-20260313T033438Z-3-001.zip/Rank_1_Year_2023_Reclassified.tif'
    modified_xml = modified_xml.replace(old_source, target_layer.source())

    modified_xml = modified_xml.replace(
        'Rank_1_Year_2023_Reclassified_tif_6cd59ee0_6bd6_4da5_a483_e006a94e9564',
        target_layer.id()
    )

    modified_xml = modified_xml.replace(
        'Rank_1_Year_2023_Reclassified.tif',
        os.path.basename(target_layer.source())
    )

    # Parse the modified XML
    doc = QDomDocument()
    doc.setContent(modified_xml)

    layout_name = f"Kiwere_Temp_{year}"

    existing = manager.layoutByName(layout_name)
    if existing:
        manager.removeLayout(existing)

    layout = QgsPrintLayout(project)
    layout.setName(layout_name)
    manager.addLayout(layout)

    context = QgsReadWriteContext()
    items, ok = layout.loadFromTemplate(doc, context)
    if not ok:
        print(f"  ERROR: Failed to load template for Rank {rank} (Year {year}).")
        manager.removeLayout(layout)
        return False

    # Configure map item
    map_item = None
    for item in layout.items():
        if hasattr(item, 'setKeepLayerSet'):
            map_item = item
            break

    if map_item:
        visible_layers = [target_layer]
        for lyr in project.mapLayers().values():
            if "Rank_" not in lyr.name() and "Reclassified" not in lyr.name():
                if lyr.isValid():
                    visible_layers.append(lyr)
        map_item.setKeepLayerSet(True)
        map_item.setLayers(visible_layers)
        map_item.refresh()

    # Update legend
    legend_item = None
    for item in layout.items():
        if hasattr(item, 'setAutoUpdateModel'):
            legend_item = item
            break

    if legend_item:
        legend_item.setAutoUpdateModel(False)
        model = legend_item.model()
        root = model.rootGroup()
        root.clear()

        # Add the target layer
        layer_node = root.addLayer(target_layer)

        # Set the top-level layer name to "Legend"
        layer_node.setName("Legend")
        layer_node.setCustomProperty("legend/title-label", "Legend")

        # Remove "Band 1: remapped (Gray)" by renaming it to "-"
        # This is the raster band sublabel stored as legend/title-label
        # For raster layers, the band name shows as a subgroup.
        # We override it using QgsMapLayerLegendUtils
        original_nodes = model.layerOriginalLegendNodes(layer_node)
        for i, node in enumerate(original_nodes):
            # Keep the original user labels (Water, Vegetation, etc.)
            # but we need to find if there's a parent label to override
            pass

        # The "Band 1: remapped (Gray)" text comes from the layer's
        # legend node-order custom property. We rename it via:
        layer_node.setCustomProperty("legend/title-label", "Legend")
        
        # Set the subgroup (band name) style to Hidden so it won't display
        # The layer node shows as "Legend", under it the band name shows
        # We need to find it in the model tree and hide it
        idx = model.node2index(layer_node)
        child_count = model.rowCount(idx)
        for i in range(child_count):
            child_idx = model.index(i, 0, idx)
            child_node = model.index2node(child_idx)
            if child_node:
                # This is the "Band 1: remapped (Gray)" tree group node
                # Rename it to "-"
                child_node.setName("-")
                if hasattr(child_node, 'setCustomProperty'):
                    child_node.setCustomProperty("legend/title-label", "-")
                # Set its legend style to Hidden to remove the label entirely
                QgsLegendRenderer.setNodeLegendStyle(child_node, QgsLegendStyle.Hidden)

        legend_item.refresh()

    # Export to PNG
    output_path = os.path.join(OUTPUT_DIR, f"Kiwere_{year}.png")
    exporter = QgsLayoutExporter(layout)

    settings = QgsLayoutExporter.ImageExportSettings()
    settings.dpi = EXPORT_DPI

    result = exporter.exportToImage(output_path, settings)

    if result == QgsLayoutExporter.Success:
        print(f"  SUCCESS: Exported {output_path}")
    else:
        print(f"  ERROR: Export failed for {output_path} (error code: {result})")

    manager.removeLayout(layout)
    return result == QgsLayoutExporter.Success


def main():
    print("=" * 60)
    print("Kiwere Batch Map Export")
    print("=" * 60)

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    if not os.path.exists(QPT_PATH):
        print(f"ERROR: QPT file not found at: {QPT_PATH}")
        return

    print(f"\nLoading template: {QPT_PATH}")
    qpt_content = load_qpt_template(QPT_PATH)

    project = QgsProject.instance()
    manager = project.layoutManager()

    print("\nAvailable raster layers in project:")
    for layer in project.mapLayers().values():
        if "Rank_" in layer.name():
            print(f"  - {layer.name()}")

    success_count = 0
    fail_count = 0

    print(f"\nProcessing {len(RANK_YEAR_MAP)} maps...\n")

    for rank in sorted(RANK_YEAR_MAP.keys()):
        year = RANK_YEAR_MAP[rank]
        print(f"[Rank {rank}] Exporting Year {year}...")

        if export_map_for_rank(rank, year, qpt_content, manager, project):
            success_count += 1
        else:
            fail_count += 1
        print()

    print("=" * 60)
    print(f"DONE: {success_count} exported, {fail_count} failed")
    print(f"Output folder: {OUTPUT_DIR}")
    print("=" * 60)


main()
