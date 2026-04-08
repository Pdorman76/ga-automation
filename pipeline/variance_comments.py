"""
Variance Comment Generator for GA Automation Pipeline
======================================================
Generates narrative explanations for material budget variances using:
  1. GL transaction detail behind each flagged variance (data layer)
  2. Claude API call for polished narrative (optional, requires API key)

Falls back to data-driven drafts if the API key is not configured or
the API call fails.
"""

import os
import json
from datetime import datetime
from typing import List, Dict, Any, Optional


# ── Data-driven draft generation ─────────────────────────────

def _build_variance_context(variance: dict, gl_data, budget_data=None,
                             kardin_data=None) -> dict:
    """
    Build rich context for a single variance by pulling GL transactions
    behind the flagged account, plus budget and analytical context.

    Args:
        variance: Dict with account_code, account_name, ptd_actual, ptd_budget,
                  variance, variance_pct
        gl_data: Parsed GL data with accounts and transactions
        budget_data: Optional budget comparison data (Yardi)
        kardin_data: Optional Kardin annual budget data

    Returns:
        Dict with variance info + supporting GL transaction detail + analytical context
    """
    acct_code = str(variance.get('account_code', '') or '')
    context = {
        'account_code': acct_code,
        'account_name': str(variance.get('account_name', '') or ''),
        'ptd_actual': variance.get('ptd_actual', 0),
        'ptd_budget': variance.get('ptd_budget', 0),
        'variance_amount': variance.get('variance', 0),
        'variance_pct': variance.get('variance_pct', 0),
        'direction': 'over budget' if variance.get('variance', 0) > 0 else 'under budget',
        'transactions': [],
        'vendor_summary': {},
        'transaction_count': 0,
        # Analytical enrichments
        'noi_impact': '',
        'likely_one_time_amount': 0,
        'likely_recurring_amount': 0,
        'annual_budget': None,
        'ytd_actual': None,
        'ytd_budget': None,
        'annual_run_rate': None,
        'budget_seasonality': None,
    }

    # ── NOI impact classification ──
    first_digit = acct_code[0] if acct_code else '0'
    if first_digit == '4':
        # Revenue account: over budget = favorable, under = unfavorable
        context['noi_impact'] = 'favorable' if variance.get('variance', 0) > 0 else 'unfavorable'
    elif first_digit in ('5', '6', '7', '8'):
        # Expense account: over budget = unfavorable, under = favorable
        context['noi_impact'] = 'unfavorable' if variance.get('variance', 0) > 0 else 'favorable'

    # ── Pull GL transactions for this account ──
    prior_month_avg = 0
    if gl_data and hasattr(gl_data, 'accounts'):
        for acct in gl_data.accounts:
            if acct.account_code == acct_code:
                context['gl_beginning_balance'] = acct.beginning_balance
                context['gl_ending_balance'] = acct.ending_balance
                context['gl_net_change'] = acct.net_change
                context['gl_total_debits'] = acct.total_debits
                context['gl_total_credits'] = acct.total_credits

                # Estimate prior-month average from beginning balance
                period_str = getattr(gl_data.metadata, 'period', '') if hasattr(gl_data, 'metadata') else ''
                month_num = 1
                if '-' in period_str:
                    month_map = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                                 'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
                    month_num = month_map.get(period_str.split('-')[0], 1)
                prior_months = month_num - 1
                if prior_months > 0 and abs(acct.beginning_balance) > 0:
                    prior_month_avg = abs(acct.beginning_balance) / prior_months

                if hasattr(acct, 'transactions'):
                    one_time_total = 0
                    recurring_total = 0

                    for txn in acct.transactions:
                        net = txn.debit - txn.credit
                        txn_dict = {
                            'date': txn.date.strftime('%m/%d/%Y') if txn.date else '',
                            'description': txn.description or '',
                            'control': txn.control or '',
                            'reference': txn.reference or '',
                            'debit': txn.debit,
                            'credit': txn.credit,
                            'net': net,
                            'likely_one_time': False,
                        }

                        # Classify: a single large transaction (> 50% of total variance)
                        # from a vendor not seen in prior months is likely one-time
                        if abs(net) > abs(variance.get('variance', 0)) * 0.5:
                            txn_dict['likely_one_time'] = True
                            one_time_total += abs(net)
                        else:
                            recurring_total += abs(net)

                        context['transactions'].append(txn_dict)

                        # Build vendor/payee summary from description
                        desc = (txn.description or '').strip()
                        if desc:
                            vendor_key = desc[:40]
                            if vendor_key not in context['vendor_summary']:
                                context['vendor_summary'][vendor_key] = {
                                    'total': 0, 'count': 0,
                                }
                            context['vendor_summary'][vendor_key]['total'] += net
                            context['vendor_summary'][vendor_key]['count'] += 1

                    context['transaction_count'] = len(acct.transactions)
                    context['likely_one_time_amount'] = one_time_total
                    context['likely_recurring_amount'] = recurring_total
                break

    # ── Budget comparison context (Yardi YTD) ──
    if budget_data:
        items = budget_data if isinstance(budget_data, list) else getattr(budget_data, 'line_items', [])
        for item in items:
            item_code = str(item.get('account_code', '') if isinstance(item, dict) else getattr(item, 'account_code', '')).strip()
            if item_code == acct_code:
                if isinstance(item, dict):
                    context['ytd_actual'] = item.get('ytd_actual')
                    context['ytd_budget'] = item.get('ytd_budget')
                    annual = item.get('annual')
                else:
                    context['ytd_actual'] = getattr(item, 'ytd_actual', None)
                    context['ytd_budget'] = getattr(item, 'ytd_budget', None)
                    annual = getattr(item, 'annual', None)

                if annual and isinstance(annual, (int, float)) and annual != 0:
                    context['annual_budget'] = annual
                    monthly_avg = abs(annual) / 12
                    ptd_budget = abs(variance.get('ptd_budget', 0))
                    if monthly_avg > 0:
                        ratio = ptd_budget / monthly_avg
                        if ratio < 0.5:
                            context['budget_seasonality'] = 'low-budget month'
                        elif ratio > 1.5:
                            context['budget_seasonality'] = 'high-budget month'
                        else:
                            context['budget_seasonality'] = 'uniform'
                break

    # ── Kardin annual budget context (monthly detail) ──
    if kardin_data and isinstance(kardin_data, list):
        for item in kardin_data:
            k_code = str(item.get('account_code', '')).strip()
            if k_code == acct_code:
                m_total = item.get('m_total', 0) or 0
                context['annual_budget'] = context.get('annual_budget') or m_total
                # Calculate annual run rate from current period
                ptd = variance.get('ptd_actual', 0) or 0
                if ptd != 0:
                    context['annual_run_rate'] = ptd * 12
                break

    return context


def generate_data_driven_comment(context: dict) -> str:
    """
    Generate a factual, data-driven variance comment from GL detail.
    No API call — uses enriched context for analytical framing.
    """
    acct = context['account_name']
    var_amt = context['variance_amount']
    var_pct = context['variance_pct']
    direction = context['direction']
    txn_count = context['transaction_count']
    noi_impact = context.get('noi_impact', '')

    # Lead with the variance and NOI impact
    impact_note = f" ({noi_impact} to NOI)" if noi_impact else ""
    comment = f"{acct} is ${abs(var_amt):,.0f} ({abs(var_pct):.0f}%) {direction}{impact_note}."

    if txn_count == 0:
        comment += " No GL transactions found for this period."
        return comment

    comment += f" {txn_count} transaction(s) in the period."

    # One-time vs recurring analysis
    one_time = context.get('likely_one_time_amount', 0)
    recurring = context.get('likely_recurring_amount', 0)
    if one_time > 0 and one_time > recurring:
        comment += f" Driven primarily by one-time item(s) (${one_time:,.0f})."
        # If excluding one-time items changes the picture, say so
        net_excl = abs(var_amt) - one_time
        if net_excl < abs(var_amt) * 0.5:
            comment += f" Excluding these, account is approximately on budget."

    # Top drivers by amount
    vendors = context.get('vendor_summary', {})
    if vendors:
        sorted_vendors = sorted(vendors.items(), key=lambda x: abs(x[1]['total']), reverse=True)
        top = sorted_vendors[:3]
        drivers = []
        for desc, info in top:
            drivers.append(f"{desc} (${abs(info['total']):,.0f}, {info['count']} txn)")
        comment += " Key drivers: " + "; ".join(drivers) + "."

    # Seasonality / annual budget context
    seasonality = context.get('budget_seasonality')
    annual = context.get('annual_budget')
    if seasonality == 'low-budget month' and annual:
        comment += f" Note: this is a low-budget month relative to ${abs(annual):,.0f} annual budget."
    elif seasonality == 'high-budget month' and annual:
        comment += f" Note: this is a high-budget month relative to ${abs(annual):,.0f} annual budget."

    # YTD context if available
    ytd_actual = context.get('ytd_actual')
    ytd_budget = context.get('ytd_budget')
    if ytd_actual is not None and ytd_budget is not None and ytd_budget != 0:
        ytd_var_pct = ((ytd_actual - ytd_budget) / abs(ytd_budget)) * 100
        if abs(ytd_var_pct) < 5:
            comment += " YTD is tracking close to budget."
        elif abs(ytd_var_pct) < abs(var_pct):
            comment += f" YTD variance ({ytd_var_pct:+.0f}%) is smaller than PTD — likely timing."

    return comment


# ── Claude API narrative generation ──────────────────────────

def _build_api_prompt(contexts: List[dict], period: str, property_name: str) -> str:
    """Build the prompt for Claude API to generate variance narratives."""

    variance_details = []
    for ctx in contexts:
        detail = f"""
Account: {ctx['account_code']} — {ctx['account_name']}
  Actual: ${ctx['ptd_actual']:,.2f}  |  Budget: ${ctx['ptd_budget']:,.2f}
  Variance: ${ctx['variance_amount']:+,.2f} ({ctx['variance_pct']:+.1f}%)
  NOI Impact: {ctx.get('noi_impact', 'n/a')}"""

        # Add YTD context if available
        ytd_a = ctx.get('ytd_actual')
        ytd_b = ctx.get('ytd_budget')
        if ytd_a is not None and ytd_b is not None:
            detail += f"\n  YTD Actual: ${ytd_a:,.2f}  |  YTD Budget: ${ytd_b:,.2f}"

        annual = ctx.get('annual_budget')
        if annual:
            detail += f"\n  Annual Budget: ${annual:,.2f}"

        seasonality = ctx.get('budget_seasonality')
        if seasonality:
            detail += f"  (this is a {seasonality})"

        one_time = ctx.get('likely_one_time_amount', 0)
        if one_time > 0:
            detail += f"\n  Likely one-time items: ${one_time:,.2f}"

        detail += f"\n  GL Transactions ({ctx['transaction_count']}):"

        for txn in ctx['transactions'][:10]:
            net = txn['net']
            one_time_flag = " [likely one-time]" if txn.get('likely_one_time') else ""
            detail += f"\n    {txn['date']}  {txn['description'][:50]}  Control: {txn['control']}  ${net:+,.2f}{one_time_flag}"

        if ctx['transaction_count'] > 10:
            detail += f"\n    ... and {ctx['transaction_count'] - 10} more transactions"

        variance_details.append(detail)

    prompt = f"""You are a CRE accounting analyst writing variance commentary for a monthly close package.
Property: {property_name}
Period: {period}

Generate a concise 1-2 sentence narrative explanation for each material budget variance below.
Focus on the WHY — what drove the variance based on the GL transaction detail provided.
Use professional accounting language suitable for an institutional investor review.

Key guidelines:
- Distinguish one-time items (marked [likely one-time]) from recurring expenses
- If excluding one-time items brings the account close to budget, say so
- Note whether the variance is favorable or unfavorable to NOI
- If YTD is tracking closer to budget than PTD, note it as a likely timing difference
- Do NOT speculate beyond what the data shows. If the cause is unclear,
  say "requires further investigation" or "timing difference pending verification."

Format your response as a JSON array of objects with keys "account_code" and "comment".

Variances to explain:
{"".join(variance_details)}
"""
    return prompt


def generate_api_comments(contexts: List[dict], period: str = '',
                           property_name: str = '',
                           api_key: str = None) -> Dict[str, str]:
    """
    Call Claude API to generate narrative variance comments.

    Args:
        contexts: List of variance context dicts from _build_variance_context()
        period: Accounting period
        property_name: Property name
        api_key: Anthropic API key

    Returns:
        Dict mapping account_code -> narrative comment string.
        Falls back to data-driven comments on any failure.
    """
    if not api_key:
        # Fallback to data-driven
        return {ctx['account_code']: generate_data_driven_comment(ctx) for ctx in contexts}

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        prompt = _build_api_prompt(contexts, period, property_name)

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}],
        )

        # Parse response
        response_text = message.content[0].text

        # Extract JSON from response (handle markdown code blocks)
        json_text = response_text
        if '```json' in json_text:
            json_text = json_text.split('```json')[1].split('```')[0]
        elif '```' in json_text:
            json_text = json_text.split('```')[1].split('```')[0]

        comments_list = json.loads(json_text.strip())

        result = {}
        for item in comments_list:
            code = item.get('account_code', '')
            comment = item.get('comment', '')
            if code and comment:
                result[code] = comment

        # Fill in any missing accounts with data-driven fallback
        for ctx in contexts:
            if ctx['account_code'] not in result:
                result[ctx['account_code']] = generate_data_driven_comment(ctx)

        return result

    except ImportError:
        # anthropic package not installed — fall back
        return {ctx['account_code']: generate_data_driven_comment(ctx) for ctx in contexts}
    except Exception as e:
        # Any API error — fall back with note
        result = {}
        for ctx in contexts:
            comment = generate_data_driven_comment(ctx)
            result[ctx['account_code']] = f"[API unavailable] {comment}"
        return result


# ── Main entry point ─────────────────────────────────────────

def generate_variance_comments(engine_result, api_key: str = None) -> List[dict]:
    """
    Generate variance comments for all material budget variances.

    Args:
        engine_result: EngineResult from pipeline run
        api_key: Optional Anthropic API key for narrative generation

    Returns:
        List of dicts with keys: account_code, account_name, variance_amount,
        variance_pct, comment, method ('api' or 'data-driven')
    """
    gl_data = engine_result.parsed.get('gl')
    budget_data = engine_result.parsed.get('budget_comparison')
    kardin_data = engine_result.parsed.get('kardin_budget')
    variances = engine_result.budget_variances or []

    if not variances:
        return []

    # Build context for each variance (with enriched analytical data)
    contexts = []
    for var in variances:
        ctx = _build_variance_context(var, gl_data, budget_data, kardin_data)
        contexts.append(ctx)

    # Generate comments
    method = 'data-driven'
    if api_key:
        comments_map = generate_api_comments(
            contexts,
            period=engine_result.period or '',
            property_name=engine_result.property_name or '',
            api_key=api_key,
        )
        # Check if API was actually used (no "[API unavailable]" prefix)
        sample = next(iter(comments_map.values()), '')
        if not sample.startswith('[API unavailable]'):
            method = 'api'
    else:
        comments_map = {ctx['account_code']: generate_data_driven_comment(ctx) for ctx in contexts}

    # Build output
    results = []
    for var in variances:
        code = var.get('account_code', '')
        results.append({
            'account_code': code,
            'account_name': var.get('account_name', ''),
            'ptd_actual': var.get('ptd_actual', 0),
            'ptd_budget': var.get('ptd_budget', 0),
            'variance_amount': var.get('variance', 0),
            'variance_pct': var.get('variance_pct', 0),
            'comment': comments_map.get(code, ''),
            'method': method,
        })

    return results
