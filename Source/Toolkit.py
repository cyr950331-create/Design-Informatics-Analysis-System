from operator import imod
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import os
import subprocess
import sys

# 添加父目录到系统路径，以便能够导入子模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入子模块
from Source.Analysis_of_Covariance_ANCOVA import ANCOVAAnalysisApp
from Source.Analytic_Hierarchy_Process_AHP_Analysis import AnalyticHierarchyProcessAHPApp
from Source.Anderson_Darling_Test import AndersonDarlingTestApp
from Source.Bartlett_Test import BartlettTestApp
from Source.Binary_Logit_Regression_Analysis import BinaryLogitRegressionAnalysisApp
from Source.Canonical_Correlation_Analysis import CanonicalCorrelationAnalysisApp
from Source.Chi_Square_Goodness_of_Fit_Test import ChiSquareGoodnessOfFitTestApp
from Source.Chi_Squared_Test import ChiSquaredTestApp
from Source.K_Means import KMeansApp
from Source.Cochrans_Q_Test import CochransQTestApp
from Source.Collinearity_Analysis_VIF import CollinearityAnalysisVIFApp
from Source.Composite_Index_Analysis import CompositeIndexAnalysisApp
from Source.Content_Validity_Analysis import ContentValidityAnalysisApp
from Source.Coupling_Coordination_Degree_Model_Analysis import CouplingCoordinationDegreeModelAnalysisApp
from Source.CRITIC_Weighting_Method_Analysis import CRITICWeightingMethodAnalysisApp
from Source.DAgostino_K_Squared_Test import DAgostinoKSquaredTestApp
from Source.Delphi_Method_Analysis import DelphiMethodAnalysisApp
from Source.DEMATEL_Analysis import DEMATELAnalysisApp
from Source.Density_Based_Clustering_Analysis import DensityBasedClusteringAnalysisApp
from Source.Descriptive_Statistics import DescriptiveStatisticsApp
from Source.Efficacy_Coefficient_Analysis import EfficacyCoefficientAnalysisApp
from Source.Entropy_Method_Analysis import EntropyMethodAnalysisApp
from Source.Exponential_Smoothing_Method_Analysis import ExponentialSmoothingMethodAnalysisApp
from Source.Factor_Analysis import FactorAnalysisApp
from Source.Friedman_Test_Analysis import FriedmanTestApp
from Source.Fuzzy_Analytic_Hierarchy_Process_FAHP_Analysis import FuzzyAnalyticHierarchyProcessFAHPApp
from Source.Gray_Prediction_Model_Analysis import GrayPredictionModelAnalysisApp
from Source.Grey_Relational_Analysis import GreyRelationalAnalysisApp
from Source.Independence_Weighting_Method_Analysis import IndependenceWeightingMethodAnalysisApp
from Source.Independent_Samples_T_Test_Analysis import IndependentSamplesTTestAnalysisApp
from Source.Information_Entropy_Weight_Method_Analysis import InformationEntropyWeightMethodAnalysisApp
from Source.Jarque_Bera_Test import JarqueBeraTestApp
from Source.KANO_Model_Analysis import KANOModelAnalysisApp
from Source.Kappa_Consistency_Test import KappaConsistencyTestApp
from Source.Kendalls_Coordination_Coefficient import KendallsCoordinationCoefficientApp
from Source.KS_Test import KSTestApp
from Source.Lasso_Regression_Analysis import LassoRegressionAnalysisApp
from Source.Levene_Test import LeveneTestApp
from Source.Lilliefors_Test import LillieforsTestApp
from Source.Mediation_Analysis import MediationAnalysisApp
from Source.Moderated_Mediation_Analysis import ModeratedMediationAnalysisApp
from Source.Moderation_Analysis import ModerationAnalysisApp
from Source.Multi_sample_ANOVA import MultiSampleANOVAApp
from Source.Multidimensional_Scaling_MDS_Analysis import MultidimensionalScalingMDSApp
from Source.Multinomial_Logit_Regression_Analysis import MultinomialLogitRegressionApp
from Source.Multiple_choice_Question_Analysis import MultipleChoiceQuestionAnalysisApp
from Source.Multivariate_Analysis_of_Variance_MANOVA import MultivariateManovaApp
from Source.Nonlinear_Regression_Analysis import NonlinearRegressionAnalysisApp
from Source.NPS_Net_Promoter_Score_Analysis import NPSNetPromoterScoreAnalysisApp
from Source.Obstacle_Degree_Model_Analysis import ObstacleDegreeModelAnalysisApp
from Source.One_Sample_ANOVA import OneSampleANOVAApp
from Source.One_Sample_t_Test_Analysis import OneSampleTTestAnalysisApp
from Source.One_Sample_Wilcoxon_Test_Analysis import OneSampleWilcoxonTestAnalysisApp
from Source.Ordered_Logit_Regression_Analysis import OrderedLogitRegressionAnalysisApp
from Source.Ordinary_Least_Squares_Linear_Regression_Analysis import OrdinaryLeastSquaresLinearRegressionAnalysisApp
from Source.Paired_t_test_Analysis import PairedTTestAnalysisApp
from Source.Paired_Sample_Wilcoxon_Test_Analysis import PairedSampleWilcoxonTestAnalysisApp
from Source.Partial_Correlation_Analysis import PartialCorrelationAnalysisApp
from Source.Partial_Least_Squares_Regression_Analysis import PartialLeastSquaresRegressionAnalysisApp
from Source.Pearson_Correlation_Analysis import PearsonCorrelationAnalysisApp
from Source.Polynomial_Regression_Analysis import PolynomialRegressionAnalysisApp
from Source.Post_hoc_Multiple_Comparisons import PostHocMultipleComparisonsApp
from Source.Price_Sensitivity_Meter_Analysis import PriceSensitivityMeterAnalysisApp
from Source.Range_Analysis import RangeAnalysisApp
from Source.Reliability_Analysis import ReliabilityAnalysisApp
from Source.Ridge_Regression_Analysis import RidgeRegressionAnalysisApp
from Source.Robust_Linear_Regression_Analysis import RobustLinearRegressionAnalysisApp
from Source.Runs_Test import RunsTestApp
from Source.Second_Order_Clustering_Analysis import SecondOrderClusteringAnalysisApp
from Source.Shapiro_Wilk_Test import ShapiroWilkTestApp
from Source.Spearman_Correlation_Analysis import SpearmanCorrelationAnalysisApp
from Source.Test_Retest_Reliability_Analysis import TestRetestReliabilityAnalysisApp
from Source.TOPSIS_Method_Analysis import TOPSISMethodAnalysisApp
from Source.Turf_Combination_Model_Analysis import TurfCombinationModelAnalysisApp
from Source.Undesirable_SBM_Model_Analysis import UndesirableSBMModelAnalysisApp
from Source.Validity_Analysis import ValidityAnalysisApp
from Source.Within_Group_Inter_Rater_Reliability_rwg_Analysis import WithinGroupInterRaterReliabilityRwgAnalysisApp

# 全局变量
canvas = None
button_frame = None

# 定义语言字典
LANGUAGES = {
    'zh': {
        'title': "工具包",
        'switch_language': "切换语言",
        'error_message': "打开脚本 {} 时出错: {}",
        'search_placeholder': "搜索分析方法"
    },
    'en': {
        'title': "Toolkit",
        'switch_language': "Switch Language",
        'error_message': "Error opening script {}: {}",
        'search_placeholder': "Search analysis methods"
    }
}

# 定义模块映射表
MODULE_MAP = {
    "Analysis of Covariance": {
        "class": ANCOVAAnalysisApp,
        "description": {
            "zh": "协方差分析",
            "en": "Analysis of Covariance"
        }
    },
    "Analytic Hierarchy Process (AHP)": {
        "class": AnalyticHierarchyProcessAHPApp,
        "description": {
            "zh": "层次分析法",
            "en": "Analytic Hierarchy Process"
        }
    },
    "Anderson-Darling Test": {
        "class": AndersonDarlingTestApp,
        "description": {
            "zh": "Anderson Darling 检验",
            "en": "Anderson Darling Test"
        }
    },
    "Bartlett Test": {
        "class": BartlettTestApp,
        "description": {
            "zh": "Bartlett 检验",
            "en": "Bartlett Test"
        }
    },
    "Binary Logit Regression Analysis": {
        "class": BinaryLogitRegressionAnalysisApp,
        "description": {
            "zh": "二元Logit回归",
            "en": "Binary Logit Regression"
        }
    },
    "Canonical Correlation Analysis": {
        "class": CanonicalCorrelationAnalysisApp,
        "description": {
            "zh": "典型相关分析",
            "en": "Canonical Correlation Analysis"
        }
    },

    "Chi-Square Goodness-of-Fit Test": {
        "class": ChiSquareGoodnessOfFitTestApp,
        "description": {
            "zh": "卡方拟合优度检验",
            "en": "Chi Square Goodness of Fit Test"
        }
    },
    "Chi-Squared Test": {
        "class": ChiSquaredTestApp,
        "description": {
            "zh": "卡方检验",
            "en": "Chi Square Test"
        }
    },

    "Cochran's Q Test": {
        "class": CochransQTestApp,
        "description": {
            "zh": "Cochran's Q 检验",
            "en": "Cochran's Q Test"
        }
    },
    "Collinearity Analysis (VIF)": {
        "class": CollinearityAnalysisVIFApp,
        "description": {
            "zh": "共线性分析",
            "en": "Collinearity Analysis"
        }
    },
    "Composite Index": {
        "class": CompositeIndexAnalysisApp,
        "description": {
            "zh": "综合指数",
            "en": "Composite Index"
        }
    },
    "Content Validity": {
        "class": ContentValidityAnalysisApp,
        "description": {
            "zh": "内容效度",
            "en": "Content Validity"
        }
    },
    "Coupling Coordination Degree Model": {
        "class": CouplingCoordinationDegreeModelAnalysisApp,
        "description": {
            "zh": "耦合协调度模型",
            "en": "Coupling Coordination Degree Model"
        }
    },
    "CRITIC Weighting Method": {
        "class": CRITICWeightingMethodAnalysisApp,
        "description": {
            "zh": "CRITIC 权重法",
            "en": "CRITIC Weighting Method"
        }
    },
    "D'Agostino's K Squared Test": {
        "class": DAgostinoKSquaredTestApp,
        "description": {
            "zh": "D'Agostino's K Squared 检验",
            "en": "D'Agostino's K Squared Test"
        }
    },
    "Delphi Method": {
        "class": DelphiMethodAnalysisApp,
        "description": {
            "zh": "德尔菲法",
            "en": "Delphi Method"
        }
    },
    "DEMATEL Analysis": {
        "class": DEMATELAnalysisApp,
        "description": {
            "zh": "DEMATEL 分析",
            "en": "DEMATEL Analysis"
        }
    },
    "Density-Based Clustering": {
        "class": DensityBasedClusteringAnalysisApp,
        "description": {
            "zh": "密度聚类",
            "en": "Density Based Clustering"
        }
    },
    "Descriptive Statistics": {
        "class": DescriptiveStatisticsApp,
        "description": {
            "zh": "描述性统计",
            "en": "Descriptive Statistics"
        }
    },
    "Efficacy Coefficient": {
        "class": EfficacyCoefficientAnalysisApp,
        "description": {
            "zh": "功效系数",
            "en": "Efficacy Coefficient"
        }
    },
    "Entropy Method": {
        "class": EntropyMethodAnalysisApp,
        "description": {
            "zh": "熵权法",
            "en": "Entropy Method"
        }
    },
    "Exponential Smoothing Method": {
        "class": ExponentialSmoothingMethodAnalysisApp,
        "description": {
            "zh": "指数平滑法",
            "en": "Exponential Smoothing Method"
        }
    },
    "Factor Analysis": {
        "class": FactorAnalysisApp,
        "description": {
            "zh": "因子分析",
            "en": "Factor Analysis"
        }
    },
    "Friedman Test": {
        "class": FriedmanTestApp,
        "description": {
            "zh": "Friedman 检验",
            "en": "Friedman Test"
        }
    },
    "Fuzzy Analytic Hierarchy Process (FAHP)": {
        "class": FuzzyAnalyticHierarchyProcessFAHPApp,
        "description": {
            "zh": "模糊层次分析法",
            "en": "Fuzzy Analytic Hierarchy Process"
        }
    },
    "Gray Prediction Model": {
        "class": GrayPredictionModelAnalysisApp,
        "description": {
            "zh": "灰色预测模型",
            "en": "Gray Prediction Model"
        }
    },
    "Grey Relational Analysis": {
        "class": GreyRelationalAnalysisApp,
        "description": {
            "zh": "灰色关联分析",
            "en": "Grey Relational Analysis"
        }
    },
    "Independence Weighting Method": {
        "class": IndependenceWeightingMethodAnalysisApp,
        "description": {
            "zh": "独立性权重法",
            "en": "Independence Weighting Method"
        }
    },
    "Independent Samples T-Test": {
        "class": IndependentSamplesTTestAnalysisApp,
        "description": {
            "zh": "独立样本 t 检验",
            "en": "Independent Samples T Test"
        }
    },
    "Information Entropy Weight Method": {
        "class": InformationEntropyWeightMethodAnalysisApp,
        "description": {
            "zh": "信息量权重法",
            "en": "Information Entropy Weight Method"
        }
    },
    "Jarque-Bera Test": {
        "class": JarqueBeraTestApp,
        "description": {
            "zh": "Jarque Bera 检验",
            "en": "JarqueBera Test"
        }
    },
    "KANO Model": {
        "class": KANOModelAnalysisApp,
        "description": {
            "zh": "KANO 模型",
            "en": "KANO Model"
        }
    },
    "Kappa Consistency Test": {
        "class": KappaConsistencyTestApp,
        "description": {
            "zh": "Kappa 一致性检验",
            "en": "Kappa Consistency Test"
        }
    },
    "Kendall's Coordination Coefficient": {
        "class": KendallsCoordinationCoefficientApp,
        "description": {
            "zh": "Kendall协调系数",
            "en": "Kendall's Coordination Coefficient"
        }
    },
    "K-Means": {
        "class": KMeansApp,
        "description": {
            "zh": "K Means",
            "en": "K Means"
        }
    },
    "KS_Test": {
        "class": KSTestApp,
        "description": {
            "zh": "KS 检验",
            "en": "KS Test"
        }
    },
    "Lasso": {
        "class": LassoRegressionAnalysisApp,
        "description": {
            "zh": "Lasso",
            "en": "Lasso"
        }
    },
    "Levene Test": {
        "class": LeveneTestApp,
        "description": {
            "zh": "Levene 检验",
            "en": "Levene Test"
        }
    },
    "Lilliefors Test": {
        "class": LillieforsTestApp,
        "description": {
            "zh": "Lilliefors 检验",
            "en": "Lilliefors Test"
        }
    },
    "Mediation": {
        "class": MediationAnalysisApp,
        "description": {
            "zh": "中介作用",
            "en": "Mediation"
        }
    },
    "Moderated Mediation": {
        "class": ModeratedMediationAnalysisApp,
        "description": {
            "zh": "调节中介作用",
            "en": "Moderated Mediation"
        }
    },
    "Moderation": {
        "class": ModerationAnalysisApp,
        "description": {
            "zh": "调节作用",
            "en": "Moderation"
        }
    },
    "Multi-sample ANOVA": {
        "class": MultiSampleANOVAApp,
        "description": {
            "zh": "多样本方差",
            "en": "Multi sample ANOVA"
        }
    },
    "Multidimensional Scaling Analysis (MDS)": {
        "class": MultidimensionalScalingMDSApp,
        "description": {
            "zh": "多维尺度分析",
            "en": "Multidimensional Scaling Analysis"
        }
    },
    "Multinomial Logit Regression": {
        "class": MultinomialLogitRegressionApp,
        "description": {
            "zh": "多分类Logit回归",
            "en": "Multinomial Logit Regression"
        }
    },
    "Multiple-choice Question of Questionnaire": {
        "class": MultipleChoiceQuestionAnalysisApp,
        "description": {
            "zh": "问卷多选题",
            "en": "Multiple choice Question of Questionnaire"
        }
    },
    "Multivariate Analysis of Variance (MANOVA)": {
        "class": MultivariateManovaApp,
        "description": {
            "zh": "多元方差分析",
            "en": "Multivariate Analysis of Variance"
        }
    },
    "Nonlinear Regression": {
        "class": NonlinearRegressionAnalysisApp,
        "description": {
            "zh": "非线性回归",
            "en": "Nonlinear Regression"
        }
    },

    "NPS Net Promoter Score": {
        "class": NPSNetPromoterScoreAnalysisApp,
        "description": {
            "zh": "NPS净推荐值",
            "en": "NPS Net Promoter Score"
        }
    },
    "Obstacle Degree Model": {
        "class": ObstacleDegreeModelAnalysisApp,
        "description": {
            "zh": "障碍度模型",
            "en": "Obstacle Degree Model"
        }
    },
    "One-sample ANOVA": {
        "class": OneSampleANOVAApp,
        "description": {
            "zh": "单样本方差",
            "en": "One Sample ANOVA"
        }
    },
    "One-Sample t-Test": {
        "class": OneSampleTTestAnalysisApp,
        "description": {
            "zh": "单样本 t 检验",
            "en": "One Sample T Test"
        }
    },
    "One-Sample Wilcoxon Test": {
        "class": OneSampleWilcoxonTestAnalysisApp,
        "description": {
            "zh": "单样本Wilcoxon检验",
            "en": "One Sample Wilcoxon Test"
        }
    },
    "Ordered Logit Regression": {
        "class": OrderedLogitRegressionAnalysisApp,
        "description": {
            "zh": "有序Logit回归",
            "en": "Ordered Logit Regression"
        }
    },
    "OLS": {
        "class": OrdinaryLeastSquaresLinearRegressionAnalysisApp,
        "description": {
            "zh": "OLS",
            "en": "OLS"
        }
    },
    "Paired t-test": {
        "class": PairedTTestAnalysisApp,
        "description": {
            "zh": "配对 t 检验",
            "en": "Paired T Test"
        }
    },
    "Paired-Sample Wilcoxon Test": {
        "class": PairedSampleWilcoxonTestAnalysisApp,
        "description": {
            "zh": "配对样本Wilcoxon检验",
            "en": "Paired-Sample Wilcoxon Test"
        }
    },
    "Partial Correlation Analysis": {
        "class": PartialCorrelationAnalysisApp,
        "description": {
            "zh": "偏相关分析",
            "en": "Partial Correlation Analysis"
        }
    },
    "PLS": {
        "class": PartialLeastSquaresRegressionAnalysisApp,
        "description": {
            "zh": "PLS",
            "en": "PLS"
        }
    },
    "Pearson Correlation Analysis": {
        "class": PearsonCorrelationAnalysisApp,
        "description": {
            "zh": "Pearson相关性分析",
            "en": "Pearson Correlation Analysis"
        }
    },
    "Polynomial Regression": {
        "class": PolynomialRegressionAnalysisApp,
        "description": {
            "zh": "多项式回归",
            "en": "Polynomial Regression"
        }
    },
    "Post-hoc Multiple Comparison": {
        "class": PostHocMultipleComparisonsApp,
        "description": {
            "zh": "事后多重比较",
            "en": "Post Hoc Multiple Comparison"
        }
    },
    "Price Sensitivity Meter (PSM)": {
        "class": PriceSensitivityMeterAnalysisApp,
        "description": {
            "zh": "价格敏感度测试模型",
            "en": "Price Sensitivity Meter"
        }
    },
    "Range Analysis": {
        "class": RangeAnalysisApp,
        "description": {
            "zh": "极差分析",
            "en": "Range Analysis"
        }
    },
    "Reliability": {
        "class": ReliabilityAnalysisApp,
        "description": {
            "zh": "信度",
            "en": "Reliability"
        }
    },
    "Ridge Regression": {
        "class": RidgeRegressionAnalysisApp,
        "description": {
            "zh": "岭回归",
            "en": "Ridge Regression"
        }
    },
    "Robust": {
        "class": RobustLinearRegressionAnalysisApp,
        "description": {
            "zh": "Robust",
            "en": "Robust"
        }
    },
    "Runs Test": {
        "class": RunsTestApp,
        "description": {
            "zh": "游程检验",
            "en": "Runs Test"
        }
    },
    "Two-Step Clustering": {
        "class": SecondOrderClusteringAnalysisApp,
        "description": {
            "zh": "二阶聚类",
            "en": "Two Step Clustering"
        }
    },
    "Shapiro Wilk Test": {
        "class": ShapiroWilkTestApp,
        "description": {
            "zh": "Shapiro Wilk 检验",
            "en": "Shapiro Wilk Test"
        }
    },
    "Spearman Correlation Analysis": {
        "class": SpearmanCorrelationAnalysisApp,
        "description": {
            "zh": "Spearman相关性分析",
            "en": "Spearman Correlation Analysis"
        }
    },
    "Test-Retest Reliability": {
        "class": TestRetestReliabilityAnalysisApp,
        "description": {
            "zh": "重测信度",
            "en": "Test Retest Reliability"
        }
    },
    "TOPSIS": {
        "class": TOPSISMethodAnalysisApp,
        "description": {
            "zh": "TOPSIS",
            "en": "TOPSIS"
        }
    },
    "Turf Combination Model": {
        "class": TurfCombinationModelAnalysisApp,
        "description": {
            "zh": "Turf组合模型",
            "en": "Turf Combination Model"
        }
    },
    "Undesirable SBM Model": {
        "class": UndesirableSBMModelAnalysisApp,
        "description": {
            "zh": "非期望SBM模型",
            "en": "Undesirable SBM Model"
        }
    },
    "Validity": {
        "class": ValidityAnalysisApp,
        "description": {
            "zh": "效度",
            "en": "Validity"
        }
    },
    "Within-Group Inter-Rater Reliability (rwg)": {
        "class": WithinGroupInterRaterReliabilityRwgAnalysisApp,
        "description": {
            "zh": "组内评分者信度",
            "en": "Within Group Inter Rater Reliability"
        }
    },
    # 可以继续添加其他模块
    # "Module Name": {
    #     "class": ModuleClass,
    #     "description": {
    #         "zh": "中文描述",
    #         "en": "English Description"
    #     }
    # },
}

def on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

def center_button_frame():
    # 更新Canvas的滚动区域
    button_frame.update_idletasks()
    canvas.config(scrollregion=canvas.bbox(ALL))

    # 计算按钮框架的宽度和高度
    button_frame_width = button_frame.winfo_width()
    button_frame_height = button_frame.winfo_height()
    canvas_width = canvas.winfo_width()
    canvas_height = canvas.winfo_height()

    # 计算水平和垂直偏移量以实现居中
    x_offset = (canvas_width - button_frame_width) // 2 if canvas_width > button_frame_width else 0
    y_offset = (canvas_height - button_frame_height) // 2 if canvas_height > button_frame_height else 0

    # 更新Canvas中窗口的位置
    canvas.coords(canvas.find_all()[0], (x_offset, y_offset))

class ToolkitApp:
    def __init__(self, root=None):
        # 当前语言
        self.current_language = 'en'

        # 获取当前脚本所在的目录
        self.project_dir = os.path.dirname(os.path.abspath(__file__))

        # 使用模块映射表替代文件扫描
        self.modules = MODULE_MAP

        # 如果没有提供root，则创建一个新窗口
        if root is None:
            self.root = ttk.Window(themename="flatly")
            self.root.withdraw()  # 先隐藏窗口，避免初始闪烁
        else:
            self.root = root
        self.root.title(LANGUAGES[self.current_language]["title"])

        # 设置窗口图标
        # 1. 获取当前文件(Toolkit.py)所在的目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # 2. 获取当前目录的上级目录（父目录）
        parent_dir = os.path.dirname(current_dir)
        # 3. 构建上级目录下 "icon" 文件夹中的 "icon.ico" 路径
        icon_path = os.path.join(parent_dir, "icon", "icon.ico")
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            print(f"图标文件未找到: {icon_path}")

        self.create_ui()
        # 显示窗口（添加这行代码）
        if root is None:
            self.root.deiconify()

    def open_module(self, module_name):
        try:
            # 使用映射表中的类创建应用实例
            module_class = self.modules[module_name]["class"]
            module_class(ttk.Toplevel(self.root))
        except Exception as e:
            self.result_label.config(text=LANGUAGES[self.current_language]['error_message'].format(module_name, e))

    def switch_language(self):
        self.current_language = 'en' if self.current_language == 'zh' else 'zh'
        self.root.title(LANGUAGES[self.current_language]["title"])
        self.language_label.config(text=LANGUAGES[self.current_language]['switch_language'])
        self.search_entry.delete(0, ttk.END)
        self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
        self.search_entry.config(foreground='gray')

        # 更新按钮文本为当前语言
        for button, module_name in zip(self.button_list, self.modules.keys()):
            button_text = self.modules[module_name]["description"][self.current_language]
            button.config(text=button_text)
            button.configure(bootstyle=PRIMARY)

    def search_scripts(self, event=None):
        keyword = self.search_entry.get().strip()
        for button, module_name in zip(self.button_list, self.modules.keys()):
            button_text = button.cget("text")
            if keyword and keyword.lower() in button_text.lower():
                button.configure(bootstyle="danger")
            else:
                button.configure(bootstyle=PRIMARY)

    def on_entry_click(self, event):
        if self.search_entry.get() == LANGUAGES[self.current_language]['search_placeholder']:
            self.search_entry.delete(0, ttk.END)
            self.search_entry.config(foreground='black')

    def on_focusout(self, event):
        if not self.search_entry.get():
            self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
            self.search_entry.config(foreground='gray')

    def create_ui(self):
        global canvas, button_frame

        # 获取屏幕的宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # 根据屏幕分辨率动态计算窗口尺寸（占屏幕的60%）
        window_width = int(screen_width * 0.6)
        window_height = int(screen_height * 0.6)

        # 限制最小窗口尺寸，避免过小
        min_width = 1000
        min_height = 600
        window_width = max(window_width, min_width)
        window_height = max(window_height, min_height)

        # 计算窗口应该放置的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(min_width, min_height)  # 添加最小尺寸限制

        # 创建一个主框架，用于居中内容
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill=BOTH)

        # 创建搜索框
        self.search_entry = ttk.Entry(main_frame)
        self.search_entry.insert(0, LANGUAGES[self.current_language]['search_placeholder'])
        self.search_entry.config(foreground='gray')
        self.search_entry.pack(pady=10, padx=10, fill=X)
        self.search_entry.bind("<KeyRelease>", self.search_scripts)
        self.search_entry.bind("<FocusIn>", self.on_entry_click)
        self.search_entry.bind("<FocusOut>", self.on_focusout)

        # 创建一个Canvas组件
        canvas = ttk.Canvas(main_frame)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        # 创建垂直滚动条
        scrollbar = ttk.Scrollbar(main_frame, command=canvas.yview)
        scrollbar.pack(side=RIGHT, fill=Y)

        # 配置Canvas的滚动条
        canvas.configure(yscrollcommand=scrollbar.set)

        # 创建一个框架来放置按钮
        button_frame = ttk.Frame(canvas)

        # 将按钮框架添加到Canvas中
        canvas.create_window((0, 0), window=button_frame, anchor=NW)

        # 存储所有按钮的列表
        self.button_list = []

        # 创建按钮
        col = 0
        row = 0
        for module_name in self.modules.keys():
            # 根据当前语言获取按钮文本
            button_text = self.modules[module_name]["description"][self.current_language]
            button = ttk.Button(button_frame, text=button_text,
                               command=lambda m=module_name: self.open_module(m),
                               bootstyle=PRIMARY)
            button.grid(row=row, column=col, padx=20, pady=5)
            self.button_list.append(button)
            col += 1
            if col == 3:
                col = 0
                row += 1

        # 初始居中按钮框架
        center_button_frame()

        # 绑定窗口大小改变事件，重新居中按钮框架
        self.root.bind("<Configure>", lambda event: center_button_frame())

        # 绑定鼠标滚轮事件
        canvas.bind_all("<MouseWheel>", on_mousewheel)

        # 创建语言切换标签
        #self.language_label = ttk.Label(self.root, text=LANGUAGES[self.current_language]['switch_language'], cursor="hand2", foreground="gray")
        #self.language_label.pack(pady=10)
        #self.language_label.bind("<Button-1>", lambda event: self.switch_language())

        # 创建结果显示标签
        self.result_label = ttk.Label(self.root, text="", justify=LEFT)
        self.result_label.pack(pady=10)

    def run(self):
        # 运行主循环
        self.root.mainloop()

# 为了向后兼容，保留原来的运行方式
def run_app():
    app = ToolkitApp()
    app.run()

if __name__ == "__main__":
    run_app()