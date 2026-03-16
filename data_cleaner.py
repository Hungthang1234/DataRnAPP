"""
Advanced Data Cleaning Module
Handles: duplicates, missing values, data quality issues
"""

import re
import pandas as pd
import numpy as np
from typing import Dict, Tuple, List


class DataCleaner:
    """Advanced data cleaning utilities"""
    
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.original_df = df.copy()
        self.report = {}
    
    def generate_quality_report(self) -> Dict:
        """Generate data quality report for each column"""
        report = {}
        
        for col in self.df.columns:
            col_data = self.df[col]
            unique_count = col_data.nunique()
            missing_count = col_data.isna().sum()
            missing_pct = (missing_count / len(self.df)) * 100
            
            report[col] = {
                "dtype": str(col_data.dtype),
                "non_null": len(col_data) - missing_count,
                "null_count": missing_count,
                "null_percent": round(missing_pct, 2),
                "unique_values": unique_count,
                "duplicates": len(col_data) - unique_count,
                "sample_values": col_data.dropna().unique()[:5].tolist() if len(col_data.dropna()) > 0 else []
            }
        
        self.report = report
        return report
    
    def remove_duplicate_rows(self, subset: List[str] = None, keep: str = "first") -> Tuple[pd.DataFrame, Dict]:
        """
        Remove duplicate rows
        
        Args:
            subset: Columns to consider for duplicates (None = all)
            keep: 'first', 'last', or False (remove all duplicates)
        
        Returns:
            Cleaned dataframe, removal report
        """
        before_count = len(self.df)
        self.df = self.df.drop_duplicates(subset=subset, keep=keep).reset_index(drop=True)
        after_count = len(self.df)
        removed = before_count - after_count
        
        return self.df, {
            "removed_rows": removed,
            "remaining_rows": after_count,
            "percent_reduced": round((removed / before_count * 100), 2) if before_count > 0 else 0
        }
    
    def handle_missing_values(self, strategy: str = "report", numeric_fill: str = "median", 
                             categorical_fill: str = "mode") -> Tuple[pd.DataFrame, Dict]:
        """
        Handle missing values with multiple strategies
        
        Args:
            strategy: 'report' (analyze), 'drop' (drop rows), 'fill' (smart fill), 'drop_cols' (drop columns)
            numeric_fill: 'mean', 'median', 'zero'
            categorical_fill: 'mode', 'unknown', 'drop'
        
        Returns:
            Cleaned dataframe, report
        """
        missing_report = {}
        
        for col in self.df.columns:
            missing_count = self.df[col].isna().sum()
            if missing_count > 0:
                missing_report[col] = {
                    "count": missing_count,
                    "percent": round((missing_count / len(self.df)) * 100, 2)
                }
        
        if strategy == "report":
            return self.df, missing_report
        
        elif strategy == "drop":
            before = len(self.df)
            self.df = self.df.dropna().reset_index(drop=True)
            after = len(self.df)
            return self.df, {
                "rows_dropped": before - after,
                "rows_remaining": after,
                "missing_columns_affected": list(missing_report.keys())
            }
        
        elif strategy == "drop_cols":
            cols_to_drop = [col for col, info in missing_report.items() if info["percent"] > 50]
            before_cols = len(self.df.columns)
            self.df = self.df.drop(columns=cols_to_drop)
            after_cols = len(self.df.columns)
            return self.df, {
                "columns_dropped": cols_to_drop,
                "columns_before": before_cols,
                "columns_after": after_cols,
                "threshold": "50% missing"
            }
        
        elif strategy == "fill":
            fill_report = {}
            for col in self.df.columns:
                if self.df[col].isna().sum() > 0:
                    if pd.api.types.is_numeric_dtype(self.df[col]):
                        if numeric_fill == "median":
                            fill_value = self.df[col].median()
                        elif numeric_fill == "mean":
                            fill_value = self.df[col].mean()
                        else:  # zero
                            fill_value = 0
                        self.df[col].fillna(fill_value, inplace=True)
                        fill_report[col] = f"filled with {numeric_fill}: {fill_value}"
                    else:
                        if categorical_fill == "mode":
                            fill_value = self.df[col].mode()[0] if len(self.df[col].mode()) > 0 else "Unknown"
                        else:  # unknown
                            fill_value = "Unknown"
                        self.df[col].fillna(fill_value, inplace=True)
                        fill_report[col] = f"filled with {categorical_fill}: {fill_value}"
            
            return self.df, {
                "missing_values_filled": len(fill_report),
                "fill_strategies": fill_report
            }
        
        return self.df, missing_report
    
    def remove_whitespace_issues(self) -> Tuple[pd.DataFrame, Dict]:
        """
        Remove extra whitespace and normalize text
        
        Returns:
            Cleaned dataframe, report
        """
        report = {}
        
        for col in self.df.select_dtypes(include=["object", "string"]).columns:
            before = self.df[col].astype(str).str.len().sum()
            
            # Strip leading/trailing spaces
            self.df[col] = self.df[col].astype(str).str.strip()
            
            # Remove extra internal spaces
            self.df[col] = self.df[col].str.replace(r'\s+', ' ', regex=True)
            
            # Remove null strings
            self.df[col] = self.df[col].replace("nan", pd.NA)
            
            after = self.df[col].astype(str).str.len().sum()
            
            report[col] = {
                "chars_before": before,
                "chars_after": after,
                "chars_removed": before - after
            }
        
        return self.df, report
    
    def standardize_column_names(self) -> Tuple[pd.DataFrame, Dict]:
        """Standardize column names to snake_case"""
        old_names = list(self.df.columns)
        
        def to_snake_case(name: str) -> str:
            name = str(name).strip()
            name = re.sub(r'[^0-9a-zA-Z]+', '_', name)
            name = re.sub(r'([a-z0-9])([A-Z])', r'\1_\2', name)
            name = re.sub(r'_+', '_', name)
            return name.lower().strip('_') or 'column'
        
        new_names = [to_snake_case(col) for col in old_names]
        self.df.columns = new_names
        
        return self.df, {
            "renamed_columns": dict(zip(old_names, new_names)),
            "total_renamed": sum(1 for old, new in zip(old_names, new_names) if old != new)
        }
    
    def detect_outliers(self, numeric_only: bool = True, method: str = "iqr") -> Dict:
        """
        Detect potential outliers using IQR or Z-score
        
        Args:
            numeric_only: Only analyze numeric columns
            method: 'iqr' or 'zscore'
        
        Returns:
            Outlier report
        """
        outlier_report = {}
        
        numeric_df = self.df.select_dtypes(include=[np.number])
        
        for col in numeric_df.columns:
            if method == "iqr":
                Q1 = numeric_df[col].quantile(0.25)
                Q3 = numeric_df[col].quantile(0.75)
                IQR = Q3 - Q1
                lower_bound = Q1 - 1.5 * IQR
                upper_bound = Q3 + 1.5 * IQR
                
                outliers = numeric_df[
                    (numeric_df[col] < lower_bound) | (numeric_df[col] > upper_bound)
                ]
                
                outlier_report[col] = {
                    "method": "IQR",
                    "outlier_count": len(outliers),
                    "outlier_percent": round((len(outliers) / len(self.df)) * 100, 2),
                    "bounds": {"lower": round(lower_bound, 2), "upper": round(upper_bound, 2)},
                    "values": outliers[col].tolist()[:10]  # First 10 outliers
                }
            
            elif method == "zscore":
                z_scores = np.abs((numeric_df[col] - numeric_df[col].mean()) / numeric_df[col].std())
                outliers = numeric_df[z_scores > 3]
                
                outlier_report[col] = {
                    "method": "Z-Score",
                    "outlier_count": len(outliers),
                    "outlier_percent": round((len(outliers) / len(self.df)) * 100, 2),
                    "threshold": 3,
                    "values": outliers[col].tolist()[:10]
                }
        
        return outlier_report
    
    def get_cleaned_data(self) -> pd.DataFrame:
        """Get the cleaned dataframe"""
        return self.df
    
    def compare_before_after(self) -> Dict:
        """Compare original vs cleaned data"""
        return {
            "original_shape": self.original_df.shape,
            "cleaned_shape": self.df.shape,
            "rows_removed": self.original_df.shape[0] - self.df.shape[0],
            "columns_removed": self.original_df.shape[1] - self.df.shape[1],
            "original_missing": int(self.original_df.isna().sum().sum()),
            "cleaned_missing": int(self.df.isna().sum().sum()),
            "missing_reduced": int(self.original_df.isna().sum().sum()) - int(self.df.isna().sum().sum())
        }
    
    def run_full_cleaning(self, remove_duplicates: bool = True, 
                         handle_missing: str = "fill",
                         remove_whitespace: bool = True,
                         standardize_names: bool = True) -> Dict:
        """
        Run complete data cleaning pipeline
        
        Returns:
            Summary report of all cleaning steps
        """
        all_reports = {}
        
        if standardize_names:
            self.df, report = self.standardize_column_names()
            all_reports["standardize_names"] = report
        
        if remove_whitespace:
            self.df, report = self.remove_whitespace_issues()
            all_reports["whitespace"] = report
        
        if remove_duplicates:
            self.df, report = self.remove_duplicate_rows()
            all_reports["duplicates"] = report
        
        if handle_missing:
            self.df, report = self.handle_missing_values(strategy=handle_missing)
            all_reports["missing_values"] = report
        
        all_reports["quality_report"] = self.generate_quality_report()
        all_reports["before_after"] = self.compare_before_after()
        
        return all_reports
