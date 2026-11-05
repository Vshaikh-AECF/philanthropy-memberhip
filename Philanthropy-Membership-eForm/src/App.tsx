import { getContext } from '@microsoft/power-apps/app'
import { useEffect, useState, type FormEvent } from 'react'
import { MembershipOrganizationService } from './generated/services/MembershipOrganizationService'
import { NavigationService } from './generated/services/NavigationService'
import { Office365UsersService } from './generated/services/Office365UsersService'
import { PhilanthropyMembershipOrgRequestService } from './generated/services/PhilanthropyMembershipOrgRequestService'
import { TableCodeService } from './generated/services/TableCodeService'
import { WorkflowLogService } from './generated/services/WorkflowLogService'
import type { MembershipOrganization } from './generated/models/MembershipOrganizationModel'
import type { PhilanthropyMembershipOrgRequest } from './generated/models/PhilanthropyMembershipOrgRequestModel'
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
  const [year, month, day] = value.slice(0, 10).split('-')
  return `${month}/${day}/${year}`
}

const getEasternDateTime = (date: Date) => {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/New_York',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(date)
  const part = (type: Intl.DateTimeFormatPartTypes) =>
    parts.find((item) => item.type === type)?.value ?? ''

  return `${part('year')}-${part('month')}-${part('day')}T${part('hour')}:${part('minute')}:${part('second')}`
}

function App() {
  const [organizations, setOrganizations] = useState<MembershipOrganization[]>([])
  const [requester, setRequester] = useState('')
  const [requesterEmail, setRequesterEmail] = useState('')
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
  const [showSubmitConfirmation, setShowSubmitConfirmation] = useState(false)
  const [contextLoaded, setContextLoaded] = useState(false)
  const [organizationsLoaded, setOrganizationsLoaded] = useState(false)
  const [loadedRequestId, setLoadedRequestId] = useState('')
  const [notFoundRequestId, setNotFoundRequestId] = useState('')
  const [savedRequest, setSavedRequest] = useState<PhilanthropyMembershipOrgRequest | null>(null)
  const [workflowLogs, setWorkflowLogs] = useState<Awaited<ReturnType<typeof WorkflowLogService.getAll>>['data']>([])
  const [loadedWorkflowRequestId, setLoadedWorkflowRequestId] = useState('')

  const isNewOrganization = selectedOrganizationId === 'new'
  const isRequestLoading = Boolean(requestId && loadedRequestId !== requestId)
  const requestNotFound = Boolean(requestId && notFoundRequestId === requestId)
  const isWorkflowLoading = Boolean(savedRequest && loadedWorkflowRequestId !== requestId)
  const isPageLoading = isSubmitting || !contextLoaded || !organizationsLoaded || isRequestLoading || isWorkflowLoading

  useEffect(() => {
    let active = true

    void getContext()
      .then((context) => {
        if (!active) return
        const hostRequestId = context.app.queryParams.requestid
        if (hostRequestId) {
          setRequestId((currentRequestId) => currentRequestId || hostRequestId)
        }
        //return Office365UsersService.MyProfile_V2('displayName')
        return Office365UsersService.MyProfile_V2()
          .then((profileResult) => {
            console.log('profileResult', profileResult)
            if (!active) return
            if (!profileResult.success) {
              throw new Error(
                profileResult.error?.message ?? 'The Office 365 Users connector returned an error.',
              )
            }
            //const name = profileResult.data.displayName
            const name = `${profileResult?.data?.givenName ?? ""} ${profileResult?.data?.surname ?? ""}`.trim()
            if (!name) {
              throw new Error('The Office 365 Users profile did not include a display name.')
            }
            if (!profileResult.data.mail) {
              throw new Error('The Office 365 Users profile did not include an email address.')
            }
            setRequester(name)
            setRequesterEmail(profileResult.data.mail)
          })
      })
      .catch((error: unknown) => {
        if (active) {
          setPageError(
            `Your signed-in name could not be loaded: ${describeError(error, 'Power Apps context is unavailable.')}`,
          )
        }
      })
      .finally(() => {
        if (active) setContextLoaded(true)
      })

    void MembershipOrganizationService.getAll()
      .then((result) => {
        if (!active) return
        if (!result.success) {
          setPageError(
            `Membership organizations could not be loaded: ${result.error?.message ?? 'The data source returned an error.'}`,
          )
          setOrganizationsLoaded(true)
          return
        }
        setOrganizations(
          [...result.data].sort((a, b) =>
            a.MemberOrganization.localeCompare(b.MemberOrganization, undefined, { sensitivity: 'base' }),
          ),
        )
        setOrganizationsLoaded(true)
      })
      .catch((error: unknown) => {
        if (active) {
          setPageError(
            `Membership organizations could not be loaded: ${describeError(error, 'The data source is unavailable.')}`,
          )
          setOrganizationsLoaded(true)
        }
      })

    return () => {
      active = false
    }
  }, [])

  useEffect(() => {
    if (!requestId || !organizationsLoaded) return

    let active = true

    const loadRequest = async () => {
      try {
        if (!/^\d+$/.test(requestId)) {
          setNotFoundRequestId(requestId)
          return
        }

        const result = await PhilanthropyMembershipOrgRequestService.getAll({
          filter: `Id eq ${Number(requestId)}`,
          top: 1,
          select: [
            'Id',
            'MembershipOrganizationId',
            'Requester',
            'Amount',
            'StartDate',
            'EndDate',
            'Description',
            'IsDeleted',
          ],
        })
        if (!active) return
        if (!result.success) {
          const error = result.error
          const status = error && 'status' in error ? error.status : undefined
          if (status === 404) {
            setNotFoundRequestId(requestId)
          } else {
            setPageError(
              `The request could not be loaded: ${error?.message ?? 'The data source returned an error.'}`,
            )
          }
          return
        }

        const request = result.data.find(
          (item) => Number(item.Id) === Number(requestId),
        )
        if (!request || request.IsDeleted === true) {
          setNotFoundRequestId(requestId)
          return
        }

        const organization = organizations.find(
          (item) => item.Id === request.MembershipOrganizationId,
        )
        setSavedRequest(request)
        setRequester(request.Requester)
        setSelectedOrganizationId(String(request.MembershipOrganizationId))
        setSponsor(organization?.InternalSponsor ?? '')
        setAmount(String(request.Amount))
        setStartDate(formatUsDate(request.StartDate))
        setEndDate(formatUsDate(request.EndDate))
        setDescription(request.Description ?? '')
      } catch (error) {
        if (active) {
          setPageError(`The request could not be loaded: ${describeError(error, 'The data source is unavailable.')}`)
        }
      } finally {
        if (active) setLoadedRequestId(requestId)
      }
    }

    void loadRequest()
    return () => {
      active = false
    }
  }, [organizations, organizationsLoaded, requestId])

  useEffect(() => {
    if (!savedRequest || !requestId) return

    let active = true
    void WorkflowLogService.getAll({
      filter: `ProcInstID eq ${Number(requestId)}`,
      orderBy: ['ActivityDate'],
    })
      .then((result) => {
        if (!active) return
        if (!result.success) {
          setPageError(
            `The workflow log could not be loaded: ${result.error?.message ?? 'The data source returned an error.'}`,
          )
          return
        }
        setWorkflowLogs(result.data)
      })
      .catch((error: unknown) => {
        if (active) {
          setPageError(`The workflow log could not be loaded: ${describeError(error, 'The data source is unavailable.')}`)
        }
      })
      .finally(() => {
        if (active) setLoadedWorkflowRequestId(requestId)
      })

    return () => {
      active = false
    }
  }, [requestId, savedRequest])

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

  const handleSubmit = async (
    event?: FormEvent<HTMLFormElement>,
    confirmed = false,
  ) => {
    event?.preventDefault()
    if (!confirmed) {
      const errors = validate()
      setFieldErrors(errors)
      if (Object.keys(errors).length > 0) return
      if (!requester || !requesterEmail) {
        setPageError('Your signed-in name or email is not available. The request was not submitted.')
        return
      }
      setShowSubmitConfirmation(true)
      return
    }

    if (!requester || !requesterEmail) {
      setPageError('Your signed-in name or email is not available. The request was not submitted.')
      return
    }

    setShowSubmitConfirmation(false)
    setPageError('')

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

      const requestTimestamp = getEasternDateTime(new Date())
      const requestResult = await PhilanthropyMembershipOrgRequestService.create({
        MembershipOrganizationId: organizationId,
        Requester: requester,
        RequesterEmail: requesterEmail,
        RequestDate: requestTimestamp,
        Amount: Number(amount),
        StartDate: parseUsDate(startDate),
        EndDate: parseUsDate(endDate),
        Description: description.trim() || undefined,
        CreatedAt: requestTimestamp,
        CreatedBy: requester,
        UpdateAt: requestTimestamp,
        UpdatedBy: requester,
        IsDeleted: false,
        DeletedAt: null,
        DeletedBy: null,
      })
      if (!requestResult.success) {
        throw new Error(
          `The request could not be submitted: ${requestResult.error?.message ?? 'The data source returned an error.'}`,
        )
      }
      if (requestResult.data.Id === undefined) {
        throw new Error('The request was saved without returning a request ID.')
      }

      const tableCodeResult = await TableCodeService.getAll({
        filter: "TableName eq 'PhilanthropyMembershipOrgRequest'",
        top: 1,
        select: ['TableCodeID', 'TableName'],
      })
      if (!tableCodeResult.success) {
        throw new Error(
          `The request was saved, but its workflow log could not be prepared: ${tableCodeResult.error?.message ?? 'The TableCode data source returned an error.'}`,
        )
      }
      const tableCodeId = tableCodeResult.data.find(
        (tableCode) => tableCode.TableName === 'PhilanthropyMembershipOrgRequest',
      )?.TableCodeID
      if (tableCodeId === undefined) {
        throw new Error('The request was saved, but no matching TableCode was found for its workflow log.')
      }

      const workflowTimestamp = getEasternDateTime(new Date())
      const workflowResult = await WorkflowLogService.create({
        ActivityDate: workflowTimestamp,
        ActivityName: 'Request',
        ActivityUser: requester,
        ActivityAction: 'Submit',
        ProcInstID: requestResult.data.Id,
        Status: 4,
        StartDate: workflowTimestamp,
        FinishDate: workflowTimestamp,
        TableCodeId: tableCodeId,
        NoteText: description.trim() || undefined,
        StatusDescription: 'Waiting for Philanthropy Membership Admins',
      })
      if (!workflowResult.success) {
        throw new Error(
          `The request was saved, but its workflow log could not be created: ${workflowResult.error?.message ?? 'The WorkflowLog data source returned an error.'}`,
        )
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

  if (isPageLoading) {
    return (
      <main className="page-loader" role="status" aria-live="polite">
        <span className="page-loader-spinner" aria-hidden="true" />
        <span>
          {isSubmitting
            ? 'Do NOT close the browser until the process is completed. You will get redirected once the process is completed.'
            : 'Please wait ...'}
        </span>
      </main>
    )
  }

  if (requestNotFound) {
    return (
      <main className="page-shell">
        <section className="form-panel not-found-panel" role="alert">
          <h1>404 - Request Not Found</h1>
          <p>The requested form does not exist or has been deleted.</p>
        </section>
      </main>
    )
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
                  disabled={Boolean(savedRequest)}
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
                    disabled={Boolean(savedRequest)}
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
                disabled={Boolean(savedRequest)}
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
              <span className="field-value">{savedRequest?.Requester || requester || 'Loading signed-in user…'}</span>
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
                disabled={Boolean(savedRequest)}
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
                  disabled={Boolean(savedRequest)}
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
                    disabled={Boolean(savedRequest)}
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
                  disabled={Boolean(savedRequest)}
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
                    disabled={Boolean(savedRequest)}
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
            disabled={Boolean(savedRequest)}
            value={description}
            onChange={(event) => setDescription(event.target.value)}
          />
        </section>

        <section className="form-panel form-actions">
          <p>Please click <strong>Submit</strong> when you are ready to submit this form.</p>
          {pageError && <p className="page-error" role="alert">{pageError}</p>}
          <div className="button-row">
            <button type="submit" disabled={isSubmitting || !requester || !requesterEmail || Boolean(savedRequest)}>
              Submit
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
      {savedRequest && (
        <section className="workflow-panel" aria-labelledby="workflow-heading">
          <h2 id="workflow-heading">Workflow Log</h2>
          <div className="workflow-table-wrap">
            <table className="workflow-table">
              <thead>
                <tr>
                  <th scope="col">Date</th>
                  <th scope="col">Task</th>
                  <th scope="col">User</th>
                  <th scope="col">Action</th>
                  <th scope="col">Note</th>
                </tr>
              </thead>
              <tbody>
                {workflowLogs.map((log) => (
                  <tr key={log.Id}>
                    <td>{new Date(log.ActivityDate).toLocaleString('en-US', {
                      month: 'numeric',
                      day: 'numeric',
                      year: 'numeric',
                      hour: 'numeric',
                      minute: '2-digit',
                    })}</td>
                    <td>{log.ActivityName}</td>
                    <td>{log.ActivityUser}</td>
                    <td>{log.ActivityAction}</td>
                    <td>{log.NoteText ?? ''}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}
      {showSubmitConfirmation && (
        <div className="confirmation-backdrop">
          <section
            className="confirmation-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="submit-confirmation-message"
          >
            <p id="submit-confirmation-message">
              Please confirm if you are ready to submit the eForm.
            </p>
            <div className="confirmation-actions">
              <button
                type="button"
                onClick={() => {
                  setShowSubmitConfirmation(false)
                  void handleSubmit(undefined, true)
                }}
              >
                Ok
              </button>
              <button
                className="cancel-button"
                type="button"
                onClick={() => setShowSubmitConfirmation(false)}
              >
                Cancel
              </button>
            </div>
          </section>
        </div>
      )}
    </main>
  )
}

export default App
