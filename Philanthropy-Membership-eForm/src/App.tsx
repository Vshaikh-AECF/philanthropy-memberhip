import { getContext } from '@microsoft/power-apps/app'
import { useEffect, useState, type FormEvent } from 'react'
import { MembershipOrganizationService } from './generated/services/MembershipOrganizationService'
import { NavigationService } from './generated/services/NavigationService'
import { PhilanthropyMembershipOrgRequestService } from './generated/services/PhilanthropyMembershipOrgRequestService'
import type { MembershipOrganization } from './generated/models/MembershipOrganizationModel'
import './App.css'

type FieldName =
  | 'organization'
  | 'newOrganization'
  | 'sponsor'
  | 'amount'
  | 'startDate'
  | 'endDate'

type FieldErrors = Partial<Record<FieldName, string>>

const describeError = (error: unknown, fallback: string) =>
  error instanceof Error ? error.message : fallback

const parseUsDate = (value: string) => {
  const match = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(value)
  if (!match) return ''

  const [, month, day, year] = match
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)))
  if (
    date.getUTCFullYear() !== Number(year) ||
    date.getUTCMonth() !== Number(month) - 1 ||
    date.getUTCDate() !== Number(day)
  ) {
    return ''
  }
  return `${year}-${month}-${day}`
}

const formatUsDate = (value: string) => {
  const [year, month, day] = value.split('-')
  return `${month}/${day}/${year}`
}

function App() {
  const [organizations, setOrganizations] = useState<MembershipOrganization[]>([])
  const [requester, setRequester] = useState('')
  const [selectedOrganizationId, setSelectedOrganizationId] = useState('')
  const [newOrganization, setNewOrganization] = useState('')
  const [sponsor, setSponsor] = useState('')
  const [amount, setAmount] = useState('')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [description, setDescription] = useState('')
  const [requestId, setRequestId] = useState(
    new URLSearchParams(window.location.search).get('requestid') ?? '',
  )
  const [fieldErrors, setFieldErrors] = useState<FieldErrors>({})
  const [pageError, setPageError] = useState('')
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isClosing, setIsClosing] = useState(false)

  const isNewOrganization = selectedOrganizationId === 'new'

  useEffect(() => {
    let active = true

    void getContext()
      .then((context) => {
        if (!active) return
        const name = context.user.fullName || context.user.userPrincipalName || ''
        if (name) {
          setRequester(name)
        } else {
          setPageError('Your signed-in name could not be loaded.')
        }
      })
      .catch((error: unknown) => {
        if (active) {
          setPageError(
            `Your signed-in name could not be loaded: ${describeError(error, 'Power Apps context is unavailable.')}`,
          )
        }
      })

    void MembershipOrganizationService.getAll()
      .then((result) => {
        if (!active) return
        if (!result.success) {
          setPageError(
            `Membership organizations could not be loaded: ${result.error?.message ?? 'The data source returned an error.'}`,
          )
          return
        }
        setOrganizations(
          [...result.data].sort((a, b) =>
            a.MemberOrganization.localeCompare(b.MemberOrganization, undefined, { sensitivity: 'base' }),
          ),
        )
      })
      .catch((error: unknown) => {
        if (active) {
          setPageError(
            `Membership organizations could not be loaded: ${describeError(error, 'The data source is unavailable.')}`,
          )
        }
      })

    return () => {
      active = false
    }
  }, [])

  const validate = (): FieldErrors => {
    const errors: FieldErrors = {}
    const numericAmount = Number(amount)

    if (!selectedOrganizationId) {
      errors.organization = 'Select a membership organization.'
    }
    if (isNewOrganization && !newOrganization.trim()) {
      errors.newOrganization = 'Enter the new organization name.'
    }
    if (!sponsor.trim()) {
      errors.sponsor = 'Enter an internal sponsor or liaison.'
    }
    if (
      !amount.trim() ||
      !/^\d+(\.\d{1,2})?$/.test(amount) ||
      !Number.isFinite(numericAmount)
    ) {
      errors.amount = 'Enter a valid amount using numbers only.'
    }
    const parsedStartDate = parseUsDate(startDate)
    const parsedEndDate = parseUsDate(endDate)
    if (!startDate) {
      errors.startDate = 'Select a start date.'
    } else if (!parsedStartDate) {
      errors.startDate = 'Enter a valid date as MM/DD/YYYY.'
    }
    if (!endDate) {
      errors.endDate = 'Select an end date.'
    } else if (!parsedEndDate) {
      errors.endDate = 'Enter a valid date as MM/DD/YYYY.'
    }
    if (parsedStartDate && parsedEndDate && parsedStartDate > parsedEndDate) {
      errors.endDate = 'End date must be on or after the start date.'
    }

    return errors
  }

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setPageError('')

    const errors = validate()
    setFieldErrors(errors)
    if (Object.keys(errors).length > 0) return

    if (!requester) {
      setPageError('Your signed-in name is not available. The request was not submitted.')
      return
    }

    setIsSubmitting(true)
    try {
      let organizationId: number
      if (isNewOrganization) {
        const organizationResult = await MembershipOrganizationService.create({
          MemberOrganization: newOrganization.trim(),
          InternalSponsor: sponsor.trim(),
          Amount: Number(amount),
        })
        if (!organizationResult.success) {
          throw new Error(
            `The new organization could not be saved: ${organizationResult.error?.message ?? 'The data source returned an error.'}`,
          )
        }
        if (organizationResult.data.Id === undefined) {
          throw new Error('The new organization was saved without returning an ID.')
        }
        organizationId = organizationResult.data.Id
      } else {
        organizationId = Number(selectedOrganizationId)
      }

      const requestResult = await PhilanthropyMembershipOrgRequestService.create({
        MembershipOrganizationId: organizationId,
        Requester: requester,
        Amount: Number(amount),
        StartDate: parseUsDate(startDate),
        EndDate: parseUsDate(endDate),
        Description: description.trim() || undefined,
      })
      if (!requestResult.success) {
        throw new Error(
          `The request could not be submitted: ${requestResult.error?.message ?? 'The data source returned an error.'}`,
        )
      }
      if (requestResult.data.Id === undefined) {
        throw new Error('The request was saved without returning a request ID.')
      }

      const navigationResult = await NavigationService.getAll({
        select: ['Title', 'NavURL'],
      })
      if (!navigationResult.success) {
        throw new Error(
          `The request was saved, but its destination could not be loaded: ${navigationResult.error?.message ?? 'The Navigation data source returned an error.'}`,
        )
      }
      const destination = navigationResult.data.find(
        (navigation) => navigation.Title === 'PhilanthropyMembership',
      )?.NavURL
      if (!destination) {
        throw new Error('The request was saved, but no NavURL was found for PhilanthropyMembership.')
      }

      const createdRequestId = String(requestResult.data.Id)
      const url = new URL(destination, window.location.href)
      url.searchParams.set('requestid', createdRequestId)
      setRequestId(createdRequestId)
      window.location.href = url.href
    } catch (error) {
      setPageError(describeError(error, 'The request could not be submitted.'))
    } finally {
      setIsSubmitting(false)
    }
  }

  const handleClose = async () => {
    setPageError('')
    setIsClosing(true)
    try {
      const result = await NavigationService.getAll({
        filter: "Title eq 'HR'",
        top: 1,
        select: ['Title', 'NavURL'],
      })
      if (!result.success) {
        throw new Error(
          result.error?.message ?? 'The Navigation data source returned an error.',
        )
      }
      const destination = result.data[0]?.NavURL
      if (!destination) {
        throw new Error('No NavURL was found for the HR navigation entry.')
      }
      window.location.assign(destination)
    } catch (error) {
      setPageError(
        `The HR page could not be opened: ${describeError(error, 'The Navigation data source is unavailable.')}`,
      )
    } finally {
      setIsClosing(false)
    }
  }

  return (
    <main className="page-shell">
      <header className="topbar">
        <img
          className="foundation-logo"
          src="/aecf-logo.png"
          alt="The Annie E. Casey Foundation"
        />
        <h1>Philanthropy Membership - Request</h1>
        <nav className="header-nav" aria-label="Main navigation">
          <span>eForms</span>
          <span>OPS</span>
          <span>Help</span>
        </nav>
      </header>

      <form className="request-form" noValidate onSubmit={handleSubmit}>
        <section className="form-panel request-details" aria-label="Request details">
          <div className="field-grid">
            <div className="field">
              <span className="field-label">Form ID</span>
              <span className="field-value">{requestId}</span>
            </div>
          </div>
        </section>

        <section className="form-panel" aria-labelledby="membership-heading">
          <h2 id="membership-heading">Membership Information</h2>
          <div className="field-grid">
            <div className="field">
              <label className="field-label" htmlFor="organization">
                Name of Philanthropic Membership Organization
                <span className="required-mark" aria-hidden="true"> *</span>
              </label>
              <div className="organization-controls">
                <select
                  id="organization"
                  value={selectedOrganizationId}
                  onChange={(event) => {
                    const organizationId = event.target.value
                    const organization = organizations.find(
                      (item) => String(item.Id) === organizationId,
                    )
                    setSelectedOrganizationId(organizationId)
                    setSponsor(organization?.InternalSponsor ?? '')
                    setAmount(organization ? String(organization.Amount) : '')
                  }}
                  aria-invalid={Boolean(fieldErrors.organization)}
                  aria-describedby={fieldErrors.organization ? 'organization-error' : undefined}
                >
                  <option value="">Select an organization</option>
                  {organizations.map((organization) => (
                    <option key={organization.Id} value={organization.Id}>
                      {organization.MemberOrganization}
                    </option>
                  ))}
                  <option value="new">New Organization</option>
                </select>
                {isNewOrganization && (
                  <input
                    className="new-organization-input"
                    aria-label="New Organization Name"
                    placeholder="New Organization Name"
                    value={newOrganization}
                    onChange={(event) => setNewOrganization(event.target.value)}
                    aria-invalid={Boolean(fieldErrors.newOrganization)}
                    aria-describedby={fieldErrors.newOrganization ? 'new-organization-error' : undefined}
                  />
                )}
                {isNewOrganization && fieldErrors.newOrganization && (
                  <span className="field-error new-organization-error" id="new-organization-error">
                    {fieldErrors.newOrganization}
                  </span>
                )}
                {fieldErrors.organization && !isNewOrganization && (
                  <span className="field-error organization-error" id="organization-error">{fieldErrors.organization}</span>
                )}
              </div>
            </div>
            <div className="field">
              <label className="field-label" htmlFor="sponsor">
                Internal Sponsor/Liaison
                <span className="required-mark" aria-hidden="true"> *</span>
              </label>
              <input
                id="sponsor"
                value={sponsor}
                onChange={(event) => setSponsor(event.target.value)}
                aria-invalid={Boolean(fieldErrors.sponsor)}
                aria-describedby={fieldErrors.sponsor ? 'sponsor-error' : undefined}
              />
              {fieldErrors.sponsor && (
                <span className="field-error" id="sponsor-error">{fieldErrors.sponsor}</span>
              )}
            </div>
            <div className="field">
              <span className="field-label">Your name</span>
              <span className="field-value">{requester || 'Loading signed-in user…'}</span>
            </div>
            <div className="field">
              <label className="field-label" htmlFor="amount">
                Amount of membership fee
                <span className="required-mark" aria-hidden="true"> *</span>
              </label>
              <input
                id="amount"
                type="number"
                min="0"
                step="0.01"
                inputMode="decimal"
                value={amount}
                onChange={(event) => setAmount(event.target.value)}
                aria-invalid={Boolean(fieldErrors.amount)}
                aria-describedby={fieldErrors.amount ? 'amount-error' : 'amount-note'}
              />
              <span className="field-note" id="amount-note">
                Please note if the amount is different than last year.
              </span>
              {fieldErrors.amount && (
                <span className="field-error amount-error" id="amount-error">{fieldErrors.amount}</span>
              )}
            </div>
          </div>
        </section>

        <section className="form-panel" aria-labelledby="period-heading">
          <h2 id="period-heading">Membership Period</h2>
          <div className="field-grid">
            <div className="field">
              <label className="field-label" htmlFor="start-date">
                Start Date
                <span className="required-mark" aria-hidden="true"> *</span>
              </label>
              <div className="date-control">
                <input
                  id="start-date"
                  type="text"
                  inputMode="numeric"
                  placeholder="Start Date"
                  maxLength={10}
                  value={startDate}
                  onChange={(event) => setStartDate(event.target.value)}
                  aria-invalid={Boolean(fieldErrors.startDate)}
                  aria-describedby={fieldErrors.startDate ? 'start-date-error' : undefined}
                />
                <span className="date-picker-wrap">
                  <svg aria-hidden="true" viewBox="0 0 20 20" fill="none">
                    <rect x="2" y="3.5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="1.5" />
                    <path d="M6 2.5v3M14 2.5v3M2.5 7h15M6 10h1M10 10h1M14 10h1M6 13h1M10 13h1M14 13h1" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
                  </svg>
                  <input
                    type="date"
                    lang="en-US"
                    value={parseUsDate(startDate)}
                    onChange={(event) => setStartDate(event.target.value ? formatUsDate(event.target.value) : '')}
                    aria-label="Choose start date"
                  />
                </span>
              </div>
              {fieldErrors.startDate && (
                <span className="field-error date-error" id="start-date-error">{fieldErrors.startDate}</span>
              )}
            </div>
            <div className="field">
              <label className="field-label" htmlFor="end-date">
                End Date
                <span className="required-mark" aria-hidden="true"> *</span>
              </label>
              <div className="date-control">
                <input
                  id="end-date"
                  type="text"
                  inputMode="numeric"
                  placeholder="End Date"
                  maxLength={10}
                  value={endDate}
                  onChange={(event) => setEndDate(event.target.value)}
                  aria-invalid={Boolean(fieldErrors.endDate)}
                  aria-describedby={fieldErrors.endDate ? 'end-date-error' : undefined}
                />
                <span className="date-picker-wrap">
                  <svg aria-hidden="true" viewBox="0 0 20 20" fill="none">
                    <rect x="2" y="3.5" width="16" height="14" rx="2" stroke="currentColor" strokeWidth="1.5" />
                    <path d="M6 2.5v3M14 2.5v3M2.5 7h15M6 10h1M10 10h1M14 10h1M6 13h1M10 13h1M14 13h1" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
                  </svg>
                  <input
                    type="date"
                    lang="en-US"
                    value={parseUsDate(endDate)}
                    onChange={(event) => setEndDate(event.target.value ? formatUsDate(event.target.value) : '')}
                    aria-label="Choose end date"
                  />
                </span>
              </div>
              {fieldErrors.endDate && (
                <span className="field-error date-error" id="end-date-error">{fieldErrors.endDate}</span>
              )}
            </div>
          </div>
        </section>

        <section className="form-panel" aria-labelledby="description-heading">
          <h2 id="description-heading">Description</h2>
          <label className="visually-hidden" htmlFor="description">Description</label>
          <input
            id="description"
            type="text"
            maxLength={250}
            value={description}
            onChange={(event) => setDescription(event.target.value)}
          />
        </section>

        <section className="form-panel form-actions">
          <p>Please click <strong>Submit</strong> when you are ready to submit this form.</p>
          {pageError && <p className="page-error" role="alert">{pageError}</p>}
          <div className="button-row">
            <button type="submit" disabled={isSubmitting || !requester || Boolean(requestId)}>
              {requestId ? 'Submitted' : isSubmitting ? 'Submitting…' : 'Submit'}
            </button>
            <button
              className="secondary-button"
              type="button"
              onClick={handleClose}
              disabled={isClosing}
            >
              {isClosing ? 'Closing…' : 'Close'}
            </button>
          </div>
        </section>
      </form>
    </main>
  )
}

export default App
