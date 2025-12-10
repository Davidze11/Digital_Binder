"""
Visualization Tools for Drone Analysis
Generate heat maps, priority zones, and visual reports
"""

import cv2
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.colors import LinearSegmentedColormap
import seaborn as sns
from pathlib import Path
from typing import Dict, List, Tuple, Optional
import json
from datetime import datetime


class DroneVisualizer:
    """
    Create visualizations for drone image analysis including:
    - Disease heat maps
    - Treatment priority zones
    - Before/after comparisons
    - Statistical dashboards
    """
    
    def __init__(self, output_dir: str = "results/drone_analysis"):
        """
        Initialize the visualizer.
        
        Args:
            output_dir: Directory to save visualization outputs
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Define color schemes
        self.severity_colors = {
            'healthy': (76, 175, 80),      # Green
            'low': (255, 235, 59),         # Yellow
            'moderate': (255, 152, 0),     # Orange
            'high': (244, 67, 54),         # Red
            'critical': (136, 14, 79)      # Dark Red
        }
        
        print(f"[+] Visualizer initialized. Output: {self.output_dir}")
    
    def create_heat_map(
        self,
        original_image_path: str,
        analysis: Dict,
        output_name: Optional[str] = None
    ) -> str:
        """
        Create a disease heat map overlay on the original drone image.
        
        Args:
            original_image_path: Path to original drone image
            analysis: Analysis dictionary from DroneImageProcessor
            output_name: Custom output filename (optional)
            
        Returns:
            Path to saved heat map image
        """
        print("\n[+] Generating disease heat map...")
        
        # Load original image
        image = cv2.imread(str(original_image_path))
        if image is None:
            raise ValueError(f"Could not load image: {original_image_path}")
        
        height, width = image.shape[:2]
        
        # Create heat map overlay
        heat_map = np.zeros((height, width), dtype=np.float32)
        
        # Fill heat map based on detections
        tile_size = analysis['metadata']['tile_size']
        
        for detection in analysis['detections']:
            x, y = detection['x'], detection['y']
            confidence = detection['confidence']
            is_diseased = 'healthy' not in detection['predicted_class'].lower()
            
            # Only show diseased areas on heat map
            if is_diseased:
                # Add Gaussian-like decay for smooth heat map
                for dy in range(-tile_size//2, tile_size//2):
                    for dx in range(-tile_size//2, tile_size//2):
                        py, px = y + tile_size//2 + dy, x + tile_size//2 + dx
                        if 0 <= py < height and 0 <= px < width:
                            distance = np.sqrt(dx**2 + dy**2) / (tile_size / 2)
                            intensity = confidence * max(0, 1 - distance)
                            heat_map[py, px] = max(heat_map[py, px], intensity)
        
        # Normalize heat map
        if heat_map.max() > 0:
            heat_map = heat_map / heat_map.max()
        
        # Apply colormap
        heat_map_colored = self._apply_heat_colormap(heat_map)
        
        # Blend with original image
        alpha = 0.5
        overlay = cv2.addWeighted(image, 1-alpha, heat_map_colored, alpha, 0)
        
        # Add legend and info
        overlay = self._add_heat_map_legend(overlay, analysis)
        
        # Save
        if output_name is None:
            output_name = f"heatmap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        
        output_path = self.output_dir / output_name
        cv2.imwrite(str(output_path), overlay)
        print(f"[+] Heat map saved: {output_path}")
        
        return str(output_path)
    
    def _apply_heat_colormap(self, heat_map: np.ndarray) -> np.ndarray:
        """Apply color gradient to heat map."""
        # Create custom colormap: transparent -> yellow -> orange -> red
        heat_map_colored = np.zeros((*heat_map.shape, 3), dtype=np.uint8)
        
        for i in range(heat_map.shape[0]):
            for j in range(heat_map.shape[1]):
                intensity = heat_map[i, j]
                if intensity < 0.01:
                    continue
                elif intensity < 0.3:
                    # Yellow
                    heat_map_colored[i, j] = [0, 255, 255]
                elif intensity < 0.6:
                    # Orange
                    heat_map_colored[i, j] = [0, 165, 255]
                else:
                    # Red
                    heat_map_colored[i, j] = [0, 0, 255]
        
        return heat_map_colored
    
    def _add_heat_map_legend(self, image: np.ndarray, analysis: Dict) -> np.ndarray:
        """Add legend and statistics to heat map."""
        height, width = image.shape[:2]
        
        # Create info panel
        panel_height = 150
        panel = np.ones((panel_height, width, 3), dtype=np.uint8) * 255
        
        # Add text
        font = cv2.FONT_HERSHEY_SIMPLEX
        font_scale = 0.6
        thickness = 2
        color = (0, 0, 0)
        
        stats = analysis['summary']
        texts = [
            f"Disease Coverage: {stats['coverage_percentage']:.1f}%",
            f"Healthy Tiles: {stats['healthy_tiles']} | Diseased Tiles: {stats['diseased_tiles']}",
            f"Total Detections: {stats['tiles_with_detections']}"
        ]
        
        y_offset = 30
        for text in texts:
            cv2.putText(panel, text, (20, y_offset), font, font_scale, color, thickness)
            y_offset += 40
        
        # Add color legend
        legend_x = width - 250
        legend_y = 20
        cv2.rectangle(panel, (legend_x, legend_y), (legend_x+30, legend_y+30), (0, 255, 255), -1)
        cv2.putText(panel, "Low Risk", (legend_x+40, legend_y+20), font, 0.5, color, 1)
        
        legend_y += 40
        cv2.rectangle(panel, (legend_x, legend_y), (legend_x+30, legend_y+30), (0, 165, 255), -1)
        cv2.putText(panel, "Medium Risk", (legend_x+40, legend_y+20), font, 0.5, color, 1)
        
        legend_y += 40
        cv2.rectangle(panel, (legend_x, legend_y), (legend_x+30, legend_y+30), (0, 0, 255), -1)
        cv2.putText(panel, "High Risk", (legend_x+40, legend_y+20), font, 0.5, color, 1)
        
        # Combine with image
        result = np.vstack([image, panel])
        return result
    
    def create_priority_zones(
        self,
        original_image_path: str,
        analysis: Dict,
        num_zones: int = 5,
        output_name: Optional[str] = None
    ) -> str:
        """
        Create treatment priority zones using clustering.
        
        Args:
            original_image_path: Path to original drone image
            analysis: Analysis dictionary
            num_zones: Number of priority zones to create
            output_name: Custom output filename
            
        Returns:
            Path to saved priority zones image
        """
        print("\n[+] Generating treatment priority zones...")
        
        from sklearn.cluster import DBSCAN
        
        # Load image
        image = cv2.imread(str(original_image_path))
        height, width = image.shape[:2]
        
        # Extract diseased tile locations
        diseased_tiles = [
            d for d in analysis['detections']
            if 'healthy' not in d['predicted_class'].lower()
        ]
        
        if not diseased_tiles:
            print("[!] No diseased areas detected. No priority zones to create.")
            return None
        
        # Prepare data for clustering
        points = np.array([[d['center_x'], d['center_y']] for d in diseased_tiles])
        confidences = np.array([d['confidence'] for d in diseased_tiles])
        
        # Perform clustering
        clustering = DBSCAN(eps=200, min_samples=3).fit(points)
        labels = clustering.labels_
        
        # Create visualization
        overlay = image.copy()
        
        # Get unique clusters (excluding noise: label -1)
        unique_labels = set(labels)
        if -1 in unique_labels:
            unique_labels.remove(-1)
        
        print(f"[+] Found {len(unique_labels)} disease clusters")
        
        # Define colors for clusters
        colors = plt.cm.rainbow(np.linspace(0, 1, len(unique_labels)))
        
        # Draw clusters and calculate priorities
        cluster_info = []
        
        for idx, label in enumerate(unique_labels):
            mask = labels == label
            cluster_points = points[mask]
            cluster_confidences = confidences[mask]
            
            # Calculate cluster statistics
            center_x = int(np.mean(cluster_points[:, 0]))
            center_y = int(np.mean(cluster_points[:, 1]))
            avg_confidence = np.mean(cluster_confidences)
            area = len(cluster_points)
            
            # Calculate priority score (higher is worse)
            priority_score = area * avg_confidence
            
            cluster_info.append({
                'id': int(label),
                'center': (center_x, center_y),
                'area_tiles': int(area),
                'avg_confidence': float(avg_confidence),
                'priority_score': float(priority_score)
            })
            
            # Draw bounding box
            min_x, min_y = cluster_points.min(axis=0)
            max_x, max_y = cluster_points.max(axis=0)
            
            color = tuple(int(c * 255) for c in colors[idx][:3])
            cv2.rectangle(overlay, (int(min_x), int(min_y)), (int(max_x), int(max_y)), color, 3)
            
            # Add label
            cv2.circle(overlay, (center_x, center_y), 10, color, -1)
            cv2.putText(overlay, f"Zone {label}", (center_x + 15, center_y),
                       cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)
        
        # Sort by priority
        cluster_info.sort(key=lambda x: x['priority_score'], reverse=True)
        
        # Add priority information panel
        overlay = self._add_priority_panel(overlay, cluster_info)
        
        # Save
        if output_name is None:
            output_name = f"priority_zones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        
        output_path = self.output_dir / output_name
        cv2.imwrite(str(output_path), overlay)
        print(f"[+] Priority zones saved: {output_path}")
        
        # Save zone data
        zones_data = {
            'zones': cluster_info,
            'timestamp': datetime.now().isoformat()
        }
        json_path = self.output_dir / output_name.replace('.jpg', '_zones.json')
        with open(json_path, 'w') as f:
            json.dump(zones_data, f, indent=2)
        
        return str(output_path)
    
    def _add_priority_panel(self, image: np.ndarray, cluster_info: List[Dict]) -> np.ndarray:
        """Add priority information panel to image."""
        height, width = image.shape[:2]
        
        panel_height = min(200, 50 + len(cluster_info) * 30)
        panel = np.ones((panel_height, width, 3), dtype=np.uint8) * 255
        
        # Title
        cv2.putText(panel, "TREATMENT PRIORITY RANKING", (20, 30),
                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 0), 2)
        
        # List priorities
        y_offset = 60
        for i, zone in enumerate(cluster_info[:5]):  # Top 5
            text = f"#{i+1} Zone {zone['id']}: {zone['area_tiles']} tiles, " \
                   f"{zone['avg_confidence']:.1%} confidence"
            cv2.putText(panel, text, (30, y_offset), cv2.FONT_HERSHEY_SIMPLEX,
                       0.5, (0, 0, 0), 1)
            y_offset += 25
        
        result = np.vstack([image, panel])
        return result
    
    def create_dashboard(
        self,
        analysis: Dict,
        output_name: Optional[str] = None
    ) -> str:
        """
        Create a comprehensive dashboard with statistics and charts.
        
        Args:
            analysis: Analysis dictionary
            output_name: Custom output filename
            
        Returns:
            Path to saved dashboard image
        """
        print("\n[+] Generating analysis dashboard...")
        
        fig, axes = plt.subplots(2, 2, figsize=(16, 12))
        fig.suptitle('Drone Image Analysis Dashboard', fontsize=20, fontweight='bold')
        
        # 1. Disease Distribution Pie Chart
        ax1 = axes[0, 0]
        disease_dist = analysis['disease_distribution']
        if disease_dist:
            labels = list(disease_dist.keys())
            sizes = list(disease_dist.values())
            
            # Simplify labels
            labels = [label.split('___')[-1][:20] for label in labels]
            
            colors = sns.color_palette('husl', len(labels))
            ax1.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
            ax1.set_title('Disease Distribution', fontsize=14, fontweight='bold')
        
        # 2. Summary Statistics
        ax2 = axes[0, 1]
        ax2.axis('off')
        
        summary = analysis['summary']
        stats_text = f"""
        FIELD SUMMARY
        {'='*40}
        
        Total Area Analyzed: {summary['total_tiles_analyzed']} tiles
        Disease Coverage: {summary['coverage_percentage']:.2f}%
        
        Healthy Areas: {summary['healthy_tiles']} tiles
        Diseased Areas: {summary['diseased_tiles']} tiles
        
        Disease Ratio: {summary['disease_ratio']:.2f}
        
        RECOMMENDATIONS
        {'='*40}
        """
        
        for i, rec in enumerate(analysis['recommendations'], 1):
            stats_text += f"\n{i}. {rec}"
        
        ax2.text(0.1, 0.9, stats_text, transform=ax2.transAxes,
                fontsize=10, verticalalignment='top',
                fontfamily='monospace',
                bbox=dict(boxstyle='round', facecolor='wheat', alpha=0.5))
        
        # 3. Top Diseases Bar Chart
        ax3 = axes[1, 0]
        if disease_dist:
            # Get diseased classes only
            diseased_only = {k: v for k, v in disease_dist.items() 
                           if 'healthy' not in k.lower()}
            
            if diseased_only:
                sorted_diseases = sorted(diseased_only.items(), 
                                       key=lambda x: x[1], reverse=True)[:10]
                diseases = [d[0].split('___')[-1][:20] for d in sorted_diseases]
                counts = [d[1] for d in sorted_diseases]
                
                bars = ax3.barh(diseases, counts, color=sns.color_palette('Reds_r', len(diseases)))
                ax3.set_xlabel('Number of Detections', fontweight='bold')
                ax3.set_title('Top Disease Detections', fontsize=14, fontweight='bold')
                ax3.invert_yaxis()
                
                # Add value labels
                for i, bar in enumerate(bars):
                    width = bar.get_width()
                    ax3.text(width, bar.get_y() + bar.get_height()/2,
                           f' {int(width)}', ha='left', va='center')
        
        # 4. Confidence Distribution
        ax4 = axes[1, 1]
        confidences = [d['confidence'] for d in analysis['detections']]
        
        if confidences:
            ax4.hist(confidences, bins=20, color='skyblue', edgecolor='black', alpha=0.7)
            ax4.axvline(np.mean(confidences), color='red', linestyle='--', 
                       linewidth=2, label=f'Mean: {np.mean(confidences):.2f}')
            ax4.set_xlabel('Confidence Score', fontweight='bold')
            ax4.set_ylabel('Frequency', fontweight='bold')
            ax4.set_title('Detection Confidence Distribution', fontsize=14, fontweight='bold')
            ax4.legend()
            ax4.grid(axis='y', alpha=0.3)
        
        plt.tight_layout()
        
        # Save
        if output_name is None:
            output_name = f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        
        output_path = self.output_dir / output_name
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"[+] Dashboard saved: {output_path}")
        return str(output_path)
    
    def create_comparison_report(
        self,
        before_image: str,
        after_image: str,
        before_analysis: Dict,
        after_analysis: Dict,
        output_name: Optional[str] = None
    ) -> str:
        """
        Create before/after treatment comparison.
        
        Args:
            before_image: Path to image before treatment
            after_image: Path to image after treatment
            before_analysis: Analysis before treatment
            after_analysis: Analysis after treatment
            output_name: Custom output filename
            
        Returns:
            Path to saved comparison report
        """
        print("\n[+] Generating before/after comparison...")
        
        fig, axes = plt.subplots(2, 2, figsize=(18, 12))
        fig.suptitle('Treatment Effectiveness Report', fontsize=20, fontweight='bold')
        
        # Load images
        img_before = cv2.imread(before_image)
        img_before = cv2.cvtColor(img_before, cv2.COLOR_BGR2RGB)
        img_after = cv2.imread(after_image)
        img_after = cv2.cvtColor(img_after, cv2.COLOR_BGR2RGB)
        
        # 1. Before image
        axes[0, 0].imshow(img_before)
        axes[0, 0].set_title(f"BEFORE Treatment\nCoverage: {before_analysis['summary']['coverage_percentage']:.1f}%",
                            fontsize=12, fontweight='bold')
        axes[0, 0].axis('off')
        
        # 2. After image
        axes[0, 1].imshow(img_after)
        axes[0, 1].set_title(f"AFTER Treatment\nCoverage: {after_analysis['summary']['coverage_percentage']:.1f}%",
                            fontsize=12, fontweight='bold')
        axes[0, 1].axis('off')
        
        # 3. Improvement metrics
        ax3 = axes[1, 0]
        
        before_sum = before_analysis['summary']
        after_sum = after_analysis['summary']
        
        improvement = {
            'Disease Coverage': [before_sum['coverage_percentage'], 
                                after_sum['coverage_percentage']],
            'Diseased Tiles': [before_sum['diseased_tiles'], 
                              after_sum['diseased_tiles']],
            'Healthy Tiles': [before_sum['healthy_tiles'], 
                             after_sum['healthy_tiles']]
        }
        
        x = np.arange(len(improvement))
        width = 0.35
        
        before_vals = [v[0] for v in improvement.values()]
        after_vals = [v[1] for v in improvement.values()]
        
        bars1 = ax3.bar(x - width/2, before_vals, width, label='Before', color='coral')
        bars2 = ax3.bar(x + width/2, after_vals, width, label='After', color='lightgreen')
        
        ax3.set_ylabel('Count / Percentage', fontweight='bold')
        ax3.set_title('Treatment Impact Metrics', fontsize=14, fontweight='bold')
        ax3.set_xticks(x)
        ax3.set_xticklabels(improvement.keys(), rotation=15, ha='right')
        ax3.legend()
        ax3.grid(axis='y', alpha=0.3)
        
        # 4. Summary statistics
        ax4 = axes[1, 1]
        ax4.axis('off')
        
        reduction_pct = ((before_sum['diseased_tiles'] - after_sum['diseased_tiles']) / 
                        max(before_sum['diseased_tiles'], 1) * 100)
        
        summary_text = f"""
        TREATMENT EFFECTIVENESS SUMMARY
        {'='*50}
        
        Disease Reduction: {reduction_pct:.1f}%
        
        Before Treatment:
          • Disease Coverage: {before_sum['coverage_percentage']:.2f}%
          • Diseased Tiles: {before_sum['diseased_tiles']}
          • Healthy Tiles: {before_sum['healthy_tiles']}
        
        After Treatment:
          • Disease Coverage: {after_sum['coverage_percentage']:.2f}%
          • Diseased Tiles: {after_sum['diseased_tiles']}
          • Healthy Tiles: {after_sum['healthy_tiles']}
        
        RESULT: {'SUCCESS' if reduction_pct > 50 else 'PARTIAL IMPROVEMENT' if reduction_pct > 0 else 'NO IMPROVEMENT'}
        """
        
        ax4.text(0.1, 0.9, summary_text, transform=ax4.transAxes,
                fontsize=11, verticalalignment='top',
                fontfamily='monospace',
                bbox=dict(boxstyle='round', 
                         facecolor='lightgreen' if reduction_pct > 50 else 'lightyellow',
                         alpha=0.7))
        
        plt.tight_layout()
        
        # Save
        if output_name is None:
            output_name = f"comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        
        output_path = self.output_dir / output_name
        plt.savefig(output_path, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"[+] Comparison report saved: {output_path}")
        return str(output_path)


if __name__ == "__main__":
    print("Drone Visualizer - Ready for use")
    print("Import this module to create visualizations from drone analysis")

