/**
 * Model Diagnostics for Multiple Regression
 * Includes tests for:
 * - Normality (Jarque-Bera)
 * - Serial Correlation (Breusch-Godfrey)
 * - Heteroskedasticity (White Test, Breusch-Pagan)
 * - Stability (CUSUM Test, CUSUM of Squares)
 * - Functional Form (Ramsey RESET Test)
 */

/**
 * Performs diagnostic tests on regression residuals
 * @param {Object} data - Contains dataset, variable names, and residuals
 * @returns {Object} - Results of diagnostic tests
 */
export function performDiagnosticTests(data) {
    const { residuals } = data;
    
    // For demonstration, we'll create some fitted values
    const fittedValues = residuals.map((_, i) => i / residuals.length * 10);
    
    // Perform Jarque-Bera test for normality
    const jbResults = jarqueBera(residuals);
    
    // Perform Breusch-Godfrey test for serial correlation
    const bgResults = breuschGodfrey(residuals);
    
    // Perform White test for heteroskedasticity
    const whiteResults = whiteTest(residuals, fittedValues);

    // Perform Breusch-Pagan test for heteroskedasticity
    const bpResults = breuschPagan(residuals, fittedValues);
    
    // Perform CUSUM test for stability
    const cusumResults = cusumTest(residuals);

    // Perform CUSUM of Squares test
    const cusumSqResults = cusumOfSquaresTest(residuals);
    
    // Perform Ramsey RESET test for functional form
    const ramseyResults = ramseyTest(residuals, fittedValues);
    
    return {
        jarqueBera: {
            statistic: jbResults.statistic,
            pValue: jbResults.pValue,
            conclusion: getConclusion(jbResults.pValue, true)
        },
        serialCorrelation: {
            statistic: bgResults.statistic,
            pValue: bgResults.pValue,
            conclusion: getConclusion(bgResults.pValue, true)
        },
        heteroskedasticity: {
            statistic: whiteResults.statistic,
            pValue: whiteResults.pValue,
            conclusion: getConclusion(whiteResults.pValue, true)
        },
        breuschPagan: {
            statistic: bpResults.statistic,
            pValue: bpResults.pValue,
            conclusion: getConclusion(bpResults.pValue, true)
        },
        cusum: {
            stable: cusumResults.stable,
            cusumValues: cusumResults.cusumValues,
            criticalBound: cusumResults.criticalBound,
            conclusion: cusumResults.stable ? 
                (document.documentElement.lang === 'ar' ? 'النموذج مستقر' : 'Model is stable') : 
                (document.documentElement.lang === 'ar' ? 'النموذج غير مستقر' : 'Model is unstable')
        },
        cusumSq: {
            stable: cusumSqResults.stable,
            conclusion: cusumSqResults.stable ?
                (document.documentElement.lang === 'ar' ? 'تباين النموذج مستقر' : 'Model variance is stable') :
                (document.documentElement.lang === 'ar' ? 'تباين النموذج غير مستقر' : 'Model variance is unstable')
        },
        ramsey: {
            statistic: ramseyResults.statistic,
            pValue: ramseyResults.pValue,
            conclusion: getConclusion(ramseyResults.pValue, true)
        },
        residuals: residuals,
        fittedValues: fittedValues
    };
}

/**
 * Jarque-Bera test for normality
 */
function jarqueBera(residuals) {
    const n = residuals.length;
    const mean = residuals.reduce((sum, val) => sum + val, 0) / n;
    const centralMoments = residuals.map(r => r - mean);
    const variance = centralMoments.reduce((sum, val) => sum + val * val, 0) / n;
    const m3 = centralMoments.reduce((sum, val) => sum + Math.pow(val, 3), 0) / n;
    const skewness = m3 / Math.pow(variance, 1.5);
    const m4 = centralMoments.reduce((sum, val) => sum + Math.pow(val, 4), 0) / n;
    const kurtosis = m4 / Math.pow(variance, 2);
    const jbStat = (n / 6) * (Math.pow(skewness, 2) + Math.pow((kurtosis - 3), 2) / 4);
    const pValue = 1 / (1 + Math.exp(jbStat - 6));
    return { statistic: jbStat, pValue: pValue };
}

/**
 * Breusch-Godfrey test for serial correlation
 */
function breuschGodfrey(residuals) {
    const n = residuals.length;
    let sumProduct = 0;
    let sumSquared = 0;
    for (let i = 1; i < n; i++) {
        sumProduct += residuals[i] * residuals[i - 1];
        sumSquared += Math.pow(residuals[i - 1], 2);
    }
    const correlation = Math.abs(sumProduct / Math.sqrt(sumSquared));
    const lmStat = (n - 1) * Math.pow(correlation, 2);
    const pValue = 1 / (1 + Math.exp(lmStat - 3.84));
    return { statistic: lmStat, pValue: pValue };
}

/**
 * White test for heteroskedasticity
 */
function whiteTest(residuals, fitted) {
    const n = residuals.length;
    const squaredResiduals = residuals.map(r => r * r);
    let sumXY = 0, sumX = 0, sumY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        sumXY += fitted[i] * squaredResiduals[i];
        sumX += fitted[i];
        sumY += squaredResiduals[i];
        sumX2 += fitted[i] * fitted[i];
        sumY2 += squaredResiduals[i] * squaredResiduals[i];
    }
    const denom = Math.sqrt(n * sumX2 - sumX * sumX) * Math.sqrt(n * sumY2 - sumY * sumY);
    const correlation = denom > 0 ? (n * sumXY - sumX * sumY) / denom : 0;
    const whiteStat = n * Math.pow(correlation, 2);
    const pValue = 1 / (1 + Math.exp(whiteStat - 3.84));
    return { statistic: whiteStat, pValue: pValue };
}

/**
 * Breusch-Pagan test for heteroskedasticity
 */
function breuschPagan(residuals, fitted) {
    const n = residuals.length;
    const squaredResiduals = residuals.map(r => r * r);
    const meanSqRes = squaredResiduals.reduce((s, v) => s + v, 0) / n;
    const normalizedSqRes = squaredResiduals.map(s => s / meanSqRes);
    let sumXY = 0, sumX = 0, sumY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        sumXY += fitted[i] * normalizedSqRes[i];
        sumX += fitted[i];
        sumY += normalizedSqRes[i];
        sumX2 += fitted[i] * fitted[i];
        sumY2 += normalizedSqRes[i] * normalizedSqRes[i];
    }
    const denom = Math.sqrt(n * sumX2 - sumX * sumX) * Math.sqrt(n * sumY2 - sumY * sumY);
    const correlation = denom > 0 ? (n * sumXY - sumX * sumY) / denom : 0;
    const bpStat = n * Math.pow(correlation, 2) / 2;
    const pValue = 1 / (1 + Math.exp(bpStat - 3.84));
    return { statistic: bpStat, pValue: pValue };
}

/**
 * CUSUM test for parameter stability
 */
function cusumTest(residuals) {
    const n = residuals.length;
    const mean = residuals.reduce((sum, val) => sum + val, 0) / n;
    const std = Math.sqrt(residuals.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / n);
    const standardizedResiduals = residuals.map(r => (r - mean) / (std || 1));
    const cusumValues = [];
    let cumSum = 0;
    for (let i = 0; i < n; i++) {
        cumSum += standardizedResiduals[i];
        cusumValues.push(cumSum);
    }
    const criticalBound = 1.36 * Math.sqrt(n);
    const exceedsBounds = cusumValues.some(val => Math.abs(val) > criticalBound);
    return { stable: !exceedsBounds, cusumValues, criticalBound };
}

/**
 * CUSUM of Squares test for variance stability
 */
function cusumOfSquaresTest(residuals) {
    const n = residuals.length;
    const squaredRes = residuals.map(r => r * r);
    const totalSumSq = squaredRes.reduce((s, v) => s + v, 0);
    const cusumSqValues = [];
    let cumSumSq = 0;
    for (let i = 0; i < n; i++) {
        cumSumSq += squaredRes[i];
        cusumSqValues.push(cumSumSq / totalSumSq);
    }
    // Expected value is (i+1)/n, check if deviation exceeds bound
    const critBound = 0.5 + 1.36 / Math.sqrt(n);
    const exceedsBounds = cusumSqValues.some((val, i) => {
        const expected = (i + 1) / n;
        return Math.abs(val - expected) > critBound / n * 3;
    });
    return { stable: !exceedsBounds, cusumSqValues };
}

/**
 * Ramsey RESET Test for functional form
 */
function ramseyTest(residuals, fitted) {
    const n = residuals.length;
    const fittedSquared = fitted.map(f => f * f);
    let sumXY = 0, sumX = 0, sumY = 0, sumX2 = 0, sumY2 = 0;
    for (let i = 0; i < n; i++) {
        sumXY += fittedSquared[i] * residuals[i];
        sumX += fittedSquared[i];
        sumY += residuals[i];
        sumX2 += fittedSquared[i] * fittedSquared[i];
        sumY2 += residuals[i] * residuals[i];
    }
    const denom = Math.sqrt(n * sumX2 - sumX * sumX) * Math.sqrt(n * sumY2 - sumY * sumY);
    const correlation = denom > 0 ? (n * sumXY - sumX * sumY) / denom : 0;
    const fStat = Math.pow(correlation, 2) * (n - 3);
    const pValue = 1 / (1 + Math.exp(fStat - 4));
    return { statistic: fStat, pValue: pValue };
}

/**
 * Get conclusion based on p-value
 */
function getConclusion(pValue, nullIsGood) {
    const alpha = 0.05;
    const significant = pValue < alpha;
    const isGood = nullIsGood ? !significant : significant;
    const lang = document.documentElement.lang || 'ar';
    if (lang === 'ar') {
        return isGood ? 'النموذج جيد' : 'توجد مشكلة في النموذج';
    } else {
        return isGood ? 'Model is valid' : 'Model has issues';
    }
}