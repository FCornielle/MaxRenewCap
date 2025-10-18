"""
Plotting and Visualization Module

This module contains functions for creating plots and visualizations
of contingency analysis and optimization results.
"""

import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from pathlib import Path


def plot_contingency_results(line_load_df, title="Contingency Analysis Results", 
                           top_n=20, figsize=(12, 8), save_path=None):
    """
    Create a bar plot of line loading results from contingency analysis.
    
    Args:
        line_load_df (pandas.DataFrame): Processed line loading data
        title (str): Plot title (default: "Contingency Analysis Results")
        top_n (int): Number of top loaded lines to show (default: 20)
        figsize (tuple): Figure size (width, height) (default: (12, 8))
        save_path (str): Path to save the plot (optional)
        
    Example:
        >>> plot_contingency_results(line_load_df, top_n=15)
    """
    try:
        # Get top N lines
        top_lines = line_load_df.head(top_n)
        
        # Create the plot
        plt.figure(figsize=figsize)
        bars = plt.bar(range(len(top_lines)), top_lines['Cargabilidad_Maxima'])
        
        # Customize the plot
        plt.xlabel('Line Index')
        plt.ylabel('Maximum Loading (%)')
        plt.title(title)
        plt.xticks(range(len(top_lines)), top_lines['Linea'], rotation=45, ha='right')
        plt.grid(True, alpha=0.3)
        
        # Add value labels on bars
        for i, bar in enumerate(bars):
            height = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f'{height:.1f}%', ha='center', va='bottom', fontsize=8)
        
        # Add horizontal line at 100% loading
        plt.axhline(y=100, color='red', linestyle='--', alpha=0.7, label='100% Loading')
        plt.legend()
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"Plot saved to '{save_path}'")
        
        plt.show()
        
    except Exception as e:
        print(f"Error creating contingency plot: {e}")


def plot_generator_optimization(results_df, title="Generator Optimization Results", 
                              figsize=(15, 10), save_path=None):
    """
    Create multiple plots for generator optimization results.
    
    Args:
        results_df (pandas.DataFrame): Optimization results DataFrame
        title (str): Plot title (default: "Generator Optimization Results")
        figsize (tuple): Figure size (width, height) (default: (15, 10))
        save_path (str): Path to save the plot (optional)
        
    Example:
        >>> plot_generator_optimization(results_df)
    """
    try:
        if results_df is None or results_df.empty:
            print("No data to plot.")
            return
        
        # Create subplots
        fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=figsize)
        fig.suptitle(title, fontsize=16)
        
        # Plot 1: Maximum Power by Substation
        substations = results_df['Subestacion']
        max_power = results_df['Potencia Maxima']
        
        ax1.bar(range(len(substations)), max_power, color='skyblue', alpha=0.7)
        ax1.set_xlabel('Substation Index')
        ax1.set_ylabel('Maximum Power (MW)')
        ax1.set_title('Maximum Safe Power by Substation')
        ax1.grid(True, alpha=0.3)
        
        # Add value labels
        for i, v in enumerate(max_power):
            ax1.text(i, v + 0.1, f'{v:.1f}', ha='center', va='bottom', fontsize=8)
        
        # Plot 2: Loading Distribution
        loading = results_df['Cargabilidad Maxima']
        ax2.hist(loading, bins=10, color='lightgreen', alpha=0.7, edgecolor='black')
        ax2.set_xlabel('Maximum Loading (%)')
        ax2.set_ylabel('Frequency')
        ax2.set_title('Distribution of Maximum Loading')
        ax2.axvline(x=100, color='red', linestyle='--', alpha=0.7, label='100% Loading')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        # Plot 3: Power vs Loading Scatter
        ax3.scatter(max_power, loading, alpha=0.7, s=60)
        ax3.set_xlabel('Maximum Power (MW)')
        ax3.set_ylabel('Maximum Loading (%)')
        ax3.set_title('Power vs Loading Relationship')
        ax3.axhline(y=100, color='red', linestyle='--', alpha=0.7, label='100% Loading')
        ax3.legend()
        ax3.grid(True, alpha=0.3)
        
        # Add correlation coefficient
        correlation = np.corrcoef(max_power, loading)[0, 1]
        ax3.text(0.05, 0.95, f'Correlation: {correlation:.3f}', 
                transform=ax3.transAxes, bbox=dict(boxstyle="round,pad=0.3", facecolor="white", alpha=0.8))
        
        # Plot 4: Critical Lines
        critical_lines = results_df['Linea Critica'].value_counts().head(10)
        ax4.bar(range(len(critical_lines)), critical_lines.values, color='orange', alpha=0.7)
        ax4.set_xlabel('Critical Line')
        ax4.set_ylabel('Frequency')
        ax4.set_title('Most Critical Lines')
        ax4.set_xticks(range(len(critical_lines)))
        ax4.set_xticklabels(critical_lines.index, rotation=45, ha='right')
        ax4.grid(True, alpha=0.3)
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"Plot saved to '{save_path}'")
        
        plt.show()
        
    except Exception as e:
        print(f"Error creating optimization plots: {e}")


def plot_power_flow_summary(bus_objects, title="Power Flow Summary", 
                          figsize=(12, 8), save_path=None):
    """
    Create a summary plot of power flow results.
    
    Args:
        bus_objects: PowerFactory bus objects
        title (str): Plot title (default: "Power Flow Summary")
        figsize (tuple): Figure size (width, height) (default: (12, 8))
        save_path (str): Path to save the plot (optional)
        
    Example:
        >>> plot_power_flow_summary(bus_objects)
    """
    try:
        # Extract voltage and power data
        bus_names = []
        voltages = []
        active_power = []
        reactive_power = []
        
        for bus in bus_objects:
            bus_names.append(bus.loc_name)
            voltages.append(bus.GetAttribute('m:u'))
            active_power.append(bus.GetAttribute('m:P'))
            reactive_power.append(bus.GetAttribute('m:Q'))
        
        # Create subplots
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=figsize)
        fig.suptitle(title, fontsize=16)
        
        # Plot 1: Voltage profile
        ax1.plot(range(len(bus_names)), voltages, 'o-', color='blue', alpha=0.7)
        ax1.set_xlabel('Bus Index')
        ax1.set_ylabel('Voltage (p.u.)')
        ax1.set_title('Bus Voltage Profile')
        ax1.axhline(y=1.0, color='red', linestyle='--', alpha=0.7, label='Nominal Voltage')
        ax1.axhline(y=0.95, color='orange', linestyle='--', alpha=0.7, label='Min Voltage (95%)')
        ax1.axhline(y=1.05, color='orange', linestyle='--', alpha=0.7, label='Max Voltage (105%)')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # Plot 2: Power profile
        ax2.plot(range(len(bus_names)), active_power, 'o-', color='green', alpha=0.7, label='Active Power')
        ax2_twin = ax2.twinx()
        ax2_twin.plot(range(len(bus_names)), reactive_power, 's-', color='red', alpha=0.7, label='Reactive Power')
        
        ax2.set_xlabel('Bus Index')
        ax2.set_ylabel('Active Power (MW)', color='green')
        ax2_twin.set_ylabel('Reactive Power (MVAr)', color='red')
        ax2.set_title('Bus Power Profile')
        ax2.grid(True, alpha=0.3)
        
        # Combine legends
        lines1, labels1 = ax2.get_legend_handles_labels()
        lines2, labels2 = ax2_twin.get_legend_handles_labels()
        ax2.legend(lines1 + lines2, labels1 + labels2, loc='upper right')
        
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            print(f"Plot saved to '{save_path}'")
        
        plt.show()
        
    except Exception as e:
        print(f"Error creating power flow plot: {e}")


def create_optimization_dashboard(results_df, line_load_df=None, 
                                output_dir='plots', figsize=(20, 12)):
    """
    Create a comprehensive dashboard with multiple plots.
    
    Args:
        results_df (pandas.DataFrame): Optimization results
        line_load_df (pandas.DataFrame): Line loading data (optional)
        output_dir (str): Output directory for plots (default: 'plots')
        figsize (tuple): Figure size (default: (20, 12))
        
    Example:
        >>> create_optimization_dashboard(results_df, line_load_df)
    """
    try:
        # Ensure output directory exists
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        # Create main dashboard
        fig = plt.figure(figsize=figsize)
        gs = fig.add_gridspec(3, 3, hspace=0.3, wspace=0.3)
        
        # Main title
        fig.suptitle('PowerFactory Optimization Dashboard', fontsize=20, fontweight='bold')
        
        # Plot 1: Power distribution (top-left, spans 2 columns)
        ax1 = fig.add_subplot(gs[0, :2])
        if results_df is not None and not results_df.empty:
            substations = results_df['Subestacion']
            max_power = results_df['Potencia Maxima']
            bars = ax1.bar(range(len(substations)), max_power, color='skyblue', alpha=0.7)
            ax1.set_title('Maximum Safe Power by Substation', fontsize=14, fontweight='bold')
            ax1.set_xlabel('Substation')
            ax1.set_ylabel('Power (MW)')
            ax1.grid(True, alpha=0.3)
            
            # Add value labels
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                        f'{height:.1f}', ha='center', va='bottom', fontsize=8)
        
        # Plot 2: Loading histogram (top-right)
        ax2 = fig.add_subplot(gs[0, 2])
        if results_df is not None and not results_df.empty:
            loading = results_df['Cargabilidad Maxima']
            ax2.hist(loading, bins=8, color='lightgreen', alpha=0.7, edgecolor='black')
            ax2.set_title('Loading Distribution', fontsize=14, fontweight='bold')
            ax2.set_xlabel('Loading (%)')
            ax2.set_ylabel('Frequency')
            ax2.axvline(x=100, color='red', linestyle='--', alpha=0.7)
            ax2.grid(True, alpha=0.3)
        
        # Plot 3: Critical lines (bottom-left, spans 2 columns)
        ax3 = fig.add_subplot(gs[1, :2])
        if results_df is not None and not results_df.empty:
            critical_lines = results_df['Linea Critica'].value_counts().head(8)
            bars = ax3.bar(range(len(critical_lines)), critical_lines.values, 
                          color='orange', alpha=0.7)
            ax3.set_title('Most Critical Lines', fontsize=14, fontweight='bold')
            ax3.set_xlabel('Critical Line')
            ax3.set_ylabel('Frequency')
            ax3.set_xticks(range(len(critical_lines)))
            ax3.set_xticklabels(critical_lines.index, rotation=45, ha='right')
            ax3.grid(True, alpha=0.3)
        
        # Plot 4: Summary statistics (bottom-right)
        ax4 = fig.add_subplot(gs[1, 2])
        if results_df is not None and not results_df.empty:
            stats_text = f"""Summary Statistics:
            
Total Substations: {len(results_df)}
Avg Max Power: {results_df['Potencia Maxima'].mean():.1f} MW
Max Power: {results_df['Potencia Maxima'].max():.1f} MW
Min Power: {results_df['Potencia Maxima'].min():.1f} MW
Avg Loading: {results_df['Cargabilidad Maxima'].mean():.1f}%
Max Loading: {results_df['Cargabilidad Maxima'].max():.1f}%"""
            
            ax4.text(0.1, 0.5, stats_text, transform=ax4.transAxes, 
                    fontsize=10, verticalalignment='center',
                    bbox=dict(boxstyle="round,pad=0.5", facecolor="lightblue", alpha=0.8))
            ax4.set_xlim(0, 1)
            ax4.set_ylim(0, 1)
            ax4.axis('off')
            ax4.set_title('Statistics', fontsize=14, fontweight='bold')
        
        # Plot 5: Line loading (if available)
        if line_load_df is not None and not line_load_df.empty:
            ax5 = fig.add_subplot(gs[2, :])
            top_lines = line_load_df.head(15)
            bars = ax5.bar(range(len(top_lines)), top_lines['Cargabilidad_Maxima'], 
                          color='lightcoral', alpha=0.7)
            ax5.set_title('Top 15 Most Loaded Lines', fontsize=14, fontweight='bold')
            ax5.set_xlabel('Line')
            ax5.set_ylabel('Loading (%)')
            ax5.set_xticks(range(len(top_lines)))
            ax5.set_xticklabels(top_lines['Linea'], rotation=45, ha='right')
            ax5.axhline(y=100, color='red', linestyle='--', alpha=0.7, label='100% Loading')
            ax5.legend()
            ax5.grid(True, alpha=0.3)
        
        # Save dashboard
        dashboard_path = Path(output_dir) / 'optimization_dashboard.png'
        plt.savefig(dashboard_path, dpi=300, bbox_inches='tight')
        print(f"Dashboard saved to '{dashboard_path}'")
        
        plt.show()
        
    except Exception as e:
        print(f"Error creating dashboard: {e}")
