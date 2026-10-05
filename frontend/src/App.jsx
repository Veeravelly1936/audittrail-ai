import { useEffect, useState } from 'react'
import {
  Activity,
  AlertOctagon,
  ArrowDown,
  ArrowUpRight,
  BadgeCheck,
  Bell,
  Check,
  ChevronDown,
  CircleDollarSign,
  Clock3,
  FileCheck2,
  FileSearch,
  FileWarning,
  MoreHorizontal,
  RefreshCw,
  Search,
  ShieldCheck,
  SlidersHorizontal,
  Sparkles,
  X,
} from 'lucide-react'

const currency = new Intl.NumberFormat('en-US', {
  style: 'currency',
  currency: 'USD',
  maximumFractionDigits: 0,
})

function App() {
  const [invoices, setInvoices] = useState([])
  const [selectedId, setSelectedId] = useState(null)
  const [query, setQuery] = useState('')
  const [filter, setFilter] = useState('All invoices')
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [notice, setNotice] = useState('')
  const [apiError, setApiError] = useState(null)
  const [notificationsOpen, setNotificationsOpen] = useState(false)

  async function loadInvoices() {
    setLoading(true)
    try {
      const response = await fetch('/api/invoices')
      if (!response.ok) throw new Error('Could not load invoices')
      const data = await response.json()
      setInvoices(data)
      setSelectedId((current) => current ?? data[0]?.id ?? null)
      setApiError(null)
    } catch {
      setApiError({
        message: 'Could not reach the invoice service. Check the backend connection.',
        retry: loadInvoices,
      })
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadInvoices()
  }, [])

  useEffect(() => {
    function handleKeyDown(event) {
      if (event.key === 'Escape') {
        setSelectedId(null)
        setNotificationsOpen(false)
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [])

  const selected = invoices.find((invoice) => invoice.id === selectedId)
  const flaggedCount = invoices.filter((invoice) => invoice.anomalies.length > 0).length
  const pendingCount = invoices.filter((invoice) => invoice.review_status === 'Pending').length
  const auditedCount = invoices.length
  const alerts = invoices
    .flatMap((invoice) => invoice.anomalies.map((anomaly, index) => ({
      ...anomaly,
      id: `${invoice.id}-${index}`,
      invoice,
    })))
    .sort((first, second) => second.invoice.invoice_date.localeCompare(first.invoice.invoice_date))
  const visibleInvoices = invoices.filter((invoice) => {
    const matchesQuery = `${invoice.vendor} ${invoice.invoice_number}`
      .toLowerCase()
      .includes(query.toLowerCase())
    const matchesFilter =
      filter === 'All invoices' ||
      (filter === 'Needs review' && invoice.review_status === 'Pending') ||
      (filter === 'Reviewed' && invoice.review_status !== 'Pending')
    return matchesQuery && matchesFilter
  })

  async function updateReview(status) {
    if (!selected || saving) return
    setSaving(true)
    setNotice('')
    try {
      const response = await fetch(`/api/invoices/${selected.id}/review`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status }),
      })
      if (!response.ok) throw new Error('Could not save review')
      const updated = await response.json()
      setInvoices((current) => current.map((invoice) => invoice.id === updated.id ? updated : invoice))
      setApiError(null)
      setNotice(`Invoice ${updated.invoice_number} marked ${status.toLowerCase()}.`)
    } catch {
      setApiError({
        message: 'Could not save this review. Check the backend connection.',
        retry: () => updateReview(status),
      })
    } finally {
      setSaving(false)
    }
  }

  function exportVisibleInvoices() {
    const columns = ['Vendor', 'Invoice Number', 'Amount', 'Date', 'Risk', 'Review Status', 'Anomalies']
    const rows = visibleInvoices.map((invoice) => [
      invoice.vendor,
      invoice.invoice_number,
      invoice.amount,
      invoice.invoice_date,
      invoice.risk_status,
      invoice.review_status,
      invoice.anomalies.map((anomaly) => anomaly.rule).join('; '),
    ])
    const csv = [columns, ...rows]
      .map((row) => row.map((value) => `"${String(value).replaceAll('"', '""')}"`).join(','))
      .join('\r\n')
    const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }))
    const link = document.createElement('a')
    link.href = url
    link.download = 'audittrail-invoices.csv'
    document.body.appendChild(link)
    link.click()
    link.remove()
    window.setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand-lockup">
          <div className="brand-mark"><Activity size={19} strokeWidth={2.4} /></div>
          <div className="brand-name">audittrail<span>AI</span></div>
        </div>
        <div className="workspace-switcher">
          <div className="workspace-avatar">VF</div>
          <div className="workspace-copy"><b>Veritas Financial</b><span>Enterprise workspace</span></div>
          <ChevronDown size={15} />
        </div>
        <div className="nav-label">WORKSPACE</div>
        <nav className="side-nav" aria-label="Main navigation">
          <a className="nav-item active" href="#dashboard"><FileSearch size={17} />Overview<span className="nav-pip" /></a>
          <a className="nav-item" href="#invoices"><FileCheck2 size={17} />Invoices<span className="nav-count">04</span></a>
          <a className="nav-item" href="#alerts"><AlertOctagon size={17} />Anomaly queue<span className="nav-count alert-count">03</span></a>
        </nav>
        <div className="nav-label rules-label">CONTROLS</div>
        <nav className="side-nav" aria-label="Controls">
          <a className="nav-item" href="#contracts"><ShieldCheck size={17} />Contract rules</a>
          <a className="nav-item" href="#activity"><Activity size={17} />Audit activity</a>
        </nav>
        <div className="sidebar-bottom">
          <div className="plan-card">
            <div className="plan-top"><span className="plan-icon"><Sparkles size={14} /></span><span>AI MONITOR</span><span className="live-dot" /></div>
            <strong>All systems operational</strong>
            <div className="plan-meter"><i /></div>
            <small>Last sync 2 minutes ago</small>
          </div>
          <div className="user-profile">
            <div className="user-avatar">JL</div>
            <div className="user-copy"><b>Jordan Lee</b><span>Senior auditor</span></div>
            <MoreHorizontal size={19} />
          </div>
        </div>
      </aside>

      <main className="main-content" id="dashboard">
        <header className="topbar">
          <div className="breadcrumb"><span>Workspace</span><b>/</b><strong>Overview</strong></div>
          <div className="topbar-actions">
            <div className="period-selector"><span className="period-dot" />September 2026<ChevronDown size={14} /></div>
            <div className="notification-wrap">
              <button
                className="icon-button"
                aria-label={`Notifications, ${alerts.length} recent alerts`}
                aria-expanded={notificationsOpen}
                aria-controls="notification-menu"
                onClick={() => setNotificationsOpen((open) => !open)}
              ><Bell size={17} />{alerts.length > 0 && <i />}</button>
              {notificationsOpen && <div className="notification-menu" id="notification-menu" role="region" aria-label="Recent audit alerts">
                <div className="notification-heading"><b>Recent audit alerts</b><span>{alerts.length}</span></div>
                {alerts.length ? alerts.slice(0, 5).map((alert) => (
                  <button
                    className="notification-item"
                    key={alert.id}
                    onClick={() => {
                      setSelectedId(alert.invoice.id)
                      setNotificationsOpen(false)
                    }}
                  >
                    <span className="notification-mark"><AlertOctagon size={14} /></span>
                    <span><b>{alert.rule}</b><small>{alert.invoice.vendor} · {alert.invoice.invoice_number}</small></span>
                  </button>
                )) : <p className="notification-empty">No recent audit alerts.</p>}
              </div>}
            </div>
            <div className="top-avatar">JL</div>
          </div>
        </header>

        <div className="page-wrap">
          <section className="page-heading">
            <div>
              <div className="eyebrow"><span />FINANCE OPERATIONS <b>·</b> SUNDAY, SEPTEMBER 27, 2026</div>
              <h1>Invoice oversight<span>.</span></h1>
              <p>Review incoming spend and resolve flagged activity.</p>
            </div>
            <button className="sync-button" onClick={loadInvoices} disabled={loading}><RefreshCw size={15} className={loading ? 'spin' : ''} /> Sync data</button>
          </section>

          {apiError && <div className="api-error-banner" role="alert" aria-live="assertive">
            <AlertOctagon size={17} />
            <span>{apiError.message}</span>
            <button onClick={apiError.retry} disabled={loading || saving}>Retry</button>
          </div>}

          <section className="summary-grid" aria-label="Invoice summary">
            <SummaryCard title="Total audited" value={auditedCount} detail="This month" icon={<BadgeCheck size={18} />} tone="green" trend="+12.4%" />
            <SummaryCard title="Flagged anomalies" value={flaggedCount} detail="Across 4 invoices" icon={<FileWarning size={18} />} tone="orange" trend="Needs attention" />
            <SummaryCard title="Pending review" value={pendingCount} detail="Awaiting your decision" icon={<Clock3 size={18} />} tone="blue" trend="View queue" />
            <div className="summary-footnote"><span className="footnote-icon"><CircleDollarSign size={17} /></span><div><b>$150/hr</b><span>Active contract rate cap</span></div><ArrowUpRight size={15} /></div>
          </section>

          <section className={`work-area ${selected ? '' : 'review-closed'}`} id="invoices">
            <div className="invoice-list-panel">
              <div className="panel-heading">
                <div><h2>Invoice register</h2><p>Recent submissions from approved vendors</p></div>
                <button className="quiet-icon" aria-label="More invoice options"><MoreHorizontal size={19} /></button>
              </div>
              <div className="table-toolbar">
                <label className="search-field"><Search size={15} /><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search vendor or invoice" /></label>
                <label className="filter-select"><SlidersHorizontal size={14} /><select value={filter} onChange={(event) => setFilter(event.target.value)}><option>All invoices</option><option>Needs review</option><option>Reviewed</option></select><ChevronDown size={13} /></label>
              </div>
              <div className="table-scroll">
                <table>
                  <thead><tr><th>VENDOR / INVOICE #</th><th>AMOUNT</th><th>DATE</th><th>RISK</th><th><span className="sr-only">Select</span></th></tr></thead>
                  <tbody>
                    {visibleInvoices.map((invoice) => (
                      <tr
                        key={invoice.id}
                        className={selectedId === invoice.id ? 'selected-row' : ''}
                        onClick={() => setSelectedId(invoice.id)}
                        tabIndex="0"
                        aria-label={`Review ${invoice.vendor}, invoice ${invoice.invoice_number}, ${invoice.risk_status} risk`}
                        aria-selected={selectedId === invoice.id}
                        onKeyDown={(event) => {
                          if (event.key === 'Enter' || event.key === ' ') {
                            event.preventDefault()
                            setSelectedId(invoice.id)
                          }
                        }}
                      >
                        <td><div className="vendor-cell"><span className={`vendor-avatar avatar-${invoice.vendor.split(' ')[0].toLowerCase()}`}>{invoice.vendor.split(' ').map((part) => part[0]).slice(0, 2).join('')}</span><span className="vendor-text"><b>{invoice.vendor}</b><small>{invoice.invoice_number}</small></span></div></td>
                        <td className="amount-cell">{currency.format(invoice.amount)}</td>
                        <td className="date-cell">{new Date(`${invoice.invoice_date}T00:00:00`).toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}</td>
                        <td><RiskBadge risk={invoice.risk_status} /></td>
                        <td className="row-arrow"><ArrowUpRight size={15} /></td>
                      </tr>
                    ))}
                    {!visibleInvoices.length && <tr><td colSpan="5" className="empty-state">No invoices match this view.</td></tr>}
                  </tbody>
                </table>
              </div>
              <div className="table-footer"><span>Showing <b>{visibleInvoices.length}</b> of <b>{invoices.length}</b> invoices</span><button onClick={exportVisibleInvoices} aria-label={`Export ${visibleInvoices.length} visible invoices as CSV`}><ArrowDown size={13} /> Export</button></div>
            </div>

            {selected && <section className="review-panel" aria-label={`Invoice review details for ${selected.invoice_number}`}>
              <div className="review-topline"><span className="section-kicker">REVIEW DETAILS</span><button className="quiet-icon" aria-label="Close invoice review" onClick={() => setSelectedId(null)}><X size={17} /></button></div>
                <div className="review-title-row"><div><h2>{selected.vendor}</h2><p>{selected.invoice_number} <span>·</span> Submitted {new Date(`${selected.invoice_date}T00:00:00`).toLocaleDateString('en-US', { month: 'long', day: 'numeric' })}</p></div><RiskBadge risk={selected.risk_status} /></div>
                <div className="invoice-total"><span>INVOICE TOTAL</span><strong>{currency.format(selected.amount)}</strong><small><span className="status-dot" />{selected.review_status === 'Pending' ? 'Awaiting review' : selected.review_status}</small></div>
                <div className="detail-section">
                  <div className="detail-heading"><h3>Line items</h3><span>CONTRACT CAP <b>$150 / HR</b></span></div>
                  <div className="line-items">
                    {selected.line_items.map((item) => {
                      const overCap = item.hourly_rate > item.contract_cap
                      return <div className={`line-item ${overCap ? 'line-item-flagged' : ''}`} key={item.description}>
                        <div className="line-item-main"><span className={`line-icon ${overCap ? 'line-icon-alert' : ''}`}>{overCap ? <AlertOctagon size={15} /> : <Check size={15} />}</span><div><b>{item.description}</b><small>{item.hours} hours</small></div></div>
                        <div className="rate-cell"><b className={overCap ? 'rate-over' : ''}>${item.hourly_rate}/hr</b><small>{overCap ? `+$${(item.hourly_rate - item.contract_cap).toFixed(0)} over cap` : 'Within contract'}</small></div>
                      </div>
                    })}
                  </div>
                </div>
                <div className={`ai-insight ${selected.anomalies.length ? 'insight-alert' : 'insight-clean'}`}>
                  <div className="insight-heading"><span className="spark-icon"><Sparkles size={14} /></span><b>AI anomaly analysis</b><span className="confidence">RULE-BASED</span></div>
                  {selected.anomalies.length ? <ul>{selected.anomalies.map((anomaly, index) => <li key={`${anomaly.rule}-${index}`}>{anomaly.explanation}</li>)}</ul> : <p>No anomalies detected. Line items are within contracted rates and no duplicate invoice number was found.</p>}
                </div>
                <div className="review-actions">
                  <div className="action-label">AUDITOR DECISION</div>
                  <div className="action-buttons">
                    <button className="decision-button approve" aria-label={`Approve invoice ${selected.invoice_number}`} onClick={() => updateReview('Approved')} disabled={saving}><Check size={15} />Approve</button>
                    <button className="decision-button flag" aria-label={`Flag invoice ${selected.invoice_number} for review`} onClick={() => updateReview('Flagged')} disabled={saving}><AlertOctagon size={15} />Flag for review</button>
                    <button className="decision-button reject" aria-label={`Reject invoice ${selected.invoice_number}`} title="Reject invoice" onClick={() => updateReview('Rejected')} disabled={saving}><X size={17} /></button>
                  </div>
                </div>
            </section>}
          </section>
          <footer className="page-footer"><span>AuditTrail AI <b>·</b> Veritas Financial Solutions</span><span><span className="footer-live" /> All systems operational <b>·</b> Rules engine v2.4</span></footer>
        </div>
      </main>
      {notice && <div className="toast" role="status"><span>{notice}</span><button onClick={() => setNotice('')} aria-label="Dismiss notification"><X size={15} /></button></div>}
    </div>
  )
}

function SummaryCard({ title, value, detail, icon, tone, trend }) {
  return <div className="summary-card"><div className="summary-card-top"><span>{title}</span><i className={`summary-icon ${tone}`}>{icon}</i></div><div className="summary-value-row"><strong>{value.toString().padStart(2, '0')}</strong><span className={`summary-trend ${tone}`}>{tone === 'green' ? <ArrowUpRight size={12} /> : null}{trend}</span></div><div className="summary-detail">{detail}</div></div>
}

function RiskBadge({ risk }) {
  return <span className={`risk-badge risk-${risk.toLowerCase()}`}><i />{risk}</span>
}

export default App