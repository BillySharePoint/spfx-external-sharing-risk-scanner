export type RiskLabel = 'Low' | 'Medium' | 'High' | 'Review Recommended' | 'Unknown';

export interface IRiskScore {
    score: number;
    label: RiskLabel;
    reasons: string[];
    recommendedActions: string[];
}
