import { getContext } from '@microsoft/power-apps/app'
import { useEffect, useState, type FormEvent } from 'react'
import { CaseyAdminFlowv1Service } from './generated/services/CaseyAdminFlowv1Service'
import { EFMModuleApproveService } from './generated/services/EFMModuleApproveService'
import { GroupDetailsService } from './generated/services/GroupDetailsService'
import { MembershipOrganizationService } from './generated/services/MembershipOrganizationService'
import { NavigationService } from './generated/services/NavigationService'
import { Office365UsersService } from './generated/services/Office365UsersService'
import { AppVersionHistoryService } from './generated/services/AppVersionHistoryService'
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

const getIndiaDateTime = (date: Date) => {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Kolkata',
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

const formatEmailDateTime = (value: string) => {
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/.exec(value)
  return match ? `${match[2]}/${match[3]}/${match[1]} ${match[4]}:${match[5]}` : value
}

const escapeHtml = (value: string) =>
  value.replace(/[&<>"']/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  })[character] ?? character)

type EmailWorkflowLog = Awaited<ReturnType<typeof WorkflowLogService.getAll>>['data'][number]

const buildRequestEmailBody = ({
  href,
  organizationName,
  requester,
  sponsor,
  amount,
  startDate,
  endDate,
  description,
  nextDestination,
  requestId,
  logs,
}: {
  href: string
  organizationName: string
  requester: string
  sponsor: string
  amount: number
  startDate: string
  endDate: string
  description: string
  nextDestination?: string
  requestId: string
  logs: EmailWorkflowLog[]
}) => `
  <p><a href="${escapeHtml(href)}">Click here to open the worklist item</a></p>
  <p style="margin:0 0 12px;color:#789900;font-weight:700;">Philanthropy Membership Org</p>
  <ol style="margin:0;padding-left:24px;">
    <li>Name of Philanthropic Membership Organization: ${escapeHtml(organizationName)}</li>
    <li>Your name and name of Internal Sponsor/Liaison if different: ${escapeHtml(requester)} / ${escapeHtml(sponsor)}</li>
    <li>Amount of membership fee (Please note if the amount is different than last year): $${escapeHtml(amount.toFixed(2))}</li>
    <li>Start and end date for membership: ${escapeHtml(startDate)} - ${escapeHtml(endDate)}</li>
    <li>In a few sentences, please explain the value of this membership and to the work of others in the Foundation. (Include this information in the Purpose field in GMS when completing the request): ${escapeHtml(description)}</li>
  </ol>
  <br>
  <p style="margin:0 0 4px;color:#789900;font-weight:700;">Process Information</p>
  ${logs.length > 0 ? `<p style="margin:0;">Task Completed: ${escapeHtml(formatEmailDateTime(logs[logs.length - 1].ActivityDate))}, ${escapeHtml(logs[logs.length - 1].ActivityUser)}, ${escapeHtml(logs[logs.length - 1].ActivityAction)}</p>` : ''}
  ${nextDestination ? `<p style="margin:0;">Next Destination: ${escapeHtml(nextDestination)}</p>` : ''}
  <p style="margin:0 0 4px;">ID: ${escapeHtml(requestId)}</p>
  <br>
  <p style="margin:0 0 4px;color:#789900;font-weight:700;">Process Log</p>
  <table style="width:100%;border-collapse:collapse;text-align:left;">
    <tbody>
      ${logs.map((log) => `<tr><td style="padding:1px 4px;">${escapeHtml(formatEmailDateTime(log.ActivityDate))}, ${escapeHtml(log.ActivityName)}, ${escapeHtml(log.ActivityUser)}, ${escapeHtml(log.ActivityAction)}, Notes: ${escapeHtml(log.NoteText ?? '')}</td></tr>`).join('')}
    </tbody>
  </table>
`

const getRequestWorkflowLogs = async (requestId: number, tableCodeId: number) => {
  const result = await WorkflowLogService.getAll({
    filter: `ProcInstID eq ${requestId}`,
  })
  if (!result.success) {
    throw new Error(`The workflow log could not be loaded: ${result.error?.message ?? 'The WorkflowLog data source returned an error.'}`)
  }
  return result.data
    .filter((log) => log.ProcInstID === requestId && log.TableCodeId === tableCodeId)
    .sort((a, b) => (a.Id ?? 0) - (b.Id ?? 0))
}

const getRequestTableCodeId = async () => {
  const result = await TableCodeService.getAll({
    filter: "TableName eq 'PhilanthropyMembershipOrgRequest'",
    top: 1,
    select: ['TableCodeID', 'TableName'],
  })
  if (!result.success) {
    throw new Error(result.error?.message ?? 'The TableCode data source returned an error.')
  }
  const tableCodeId = result.data.find(
    (tableCode) => tableCode.TableName === 'PhilanthropyMembershipOrgRequest',
  )?.TableCodeID
  if (tableCodeId === undefined) {
    throw new Error('No matching TableCode was found for PhilanthropyMembershipOrgRequest.')
  }
  return tableCodeId
}

function App() {
  const [organizations, setOrganizations] = useState<MembershipOrganization[]>([])
  const [databaseVersion, setDatabaseVersion] = useState<string | null>(null)
  const [requester, setRequester] = useState('')
  const [signedInFullName, setSignedInFullName] = useState('')
  const [requesterEmail, setRequesterEmail] = useState('')
  const [selectedOrganizationId, setSelectedOrganizationId] = useState('')
  const [newOrganization, setNewOrganization] = useState('')
  const [sponsor, setSponsor] = useState('')
  const [amount, setAmount] = useState('')
  const [approvalAction, setApprovalAction] = useState('Process')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [membershipDescription, setMembershipDescription] = useState('')
  const [description, setDescription] = useState('')
  const [requestId, setRequestId] = useState(
    new URLSearchParams(window.location.search).get('requestid') ?? '',
  )
  const [fieldErrors, setFieldErrors] = useState<FieldErrors>({})
  const [notesError, setNotesError] = useState('')
  const [pageError, setPageError] = useState('')
  const [isSubmitting, setIsSubmitting] = useState(false)
  const [isClosing, setIsClosing] = useState(false)
  const [showSubmitConfirmation, setShowSubmitConfirmation] = useState(false)
  const [showNavigationConfirmation, setShowNavigationConfirmation] = useState(false)
  const [contextLoaded, setContextLoaded] = useState(false)
  const [organizationsLoaded, setOrganizationsLoaded] = useState(false)
  const [loadedRequestId, setLoadedRequestId] = useState('')
  const [notFoundRequestId, setNotFoundRequestId] = useState('')
  const [requestAccessError, setRequestAccessError] = useState('')
  const [homeNavigationError, setHomeNavigationError] = useState('')
  const [isApprover, setIsApprover] = useState(false)
  const [savedRequest, setSavedRequest] = useState<PhilanthropyMembershipOrgRequest | null>(null)
  const [workflowLogs, setWorkflowLogs] = useState<Awaited<ReturnType<typeof WorkflowLogService.getAll>>['data']>([])
  const [loadedWorkflowRequestId, setLoadedWorkflowRequestId] = useState('')

  const isNewOrganization = selectedOrganizationId === 'new'
  const isReturnedRequester = Boolean(
    savedRequest?.Return &&
    savedRequest.RequesterEmail.trim().toLowerCase() === requesterEmail.trim().toLowerCase(),
  )
  const isDeclined = Boolean(savedRequest?.Decline)
  const isProcessed = workflowLogs[workflowLogs.length - 1]?.ActivityAction === 'Process'
  const isRequestLoading = Boolean(requestId && loadedRequestId !== requestId)
  const requestNotFound = Boolean(requestId && notFoundRequestId === requestId)
  const requestAccessFailed = Boolean(requestId && loadedRequestId === requestId && requestAccessError)
  const isWorkflowLoading = Boolean(savedRequest && loadedWorkflowRequestId !== requestId)
  const isPageLoading = isSubmitting || !contextLoaded || !organizationsLoaded || isRequestLoading || isWorkflowLoading

  useEffect(() => {
    let active = true
    let checking = false
    const reloadGuardKey = 'philanthropy-app-version-reload'

    const checkVersion = async () => {
      if (checking) return
      checking = true

      try {
        const result = await AppVersionHistoryService.getAll({
          filter: "AppName eq 'philanthropy'",
          top: 1,
          select: ['AppName', 'Version'],
        })
        if (!active) return
        if (!result.success) {
          console.error('The application version could not be loaded:', result.error)
          return
        }

        const databaseVersion = result.data.find(
          (record) => record.AppName === 'philanthropy',
        )?.Version?.trim()
        if (!databaseVersion) {
          console.error('No application version was found for AppName "philanthropy".')
          return
        }

        setDatabaseVersion(databaseVersion)
        if (databaseVersion === __APP_VERSION__) {
          sessionStorage.removeItem(reloadGuardKey)
          return
        }

        const reloadGuardValue = `${__APP_VERSION__}:${databaseVersion}`
        if (sessionStorage.getItem(reloadGuardKey) === reloadGuardValue) return
        sessionStorage.setItem(reloadGuardKey, reloadGuardValue)

        const refreshUrl = new URL(window.location.href)
        refreshUrl.searchParams.set('_appVersionRefresh', String(Date.now()))
        window.location.replace(refreshUrl.href)
      } catch (error) {
        if (active) {
          console.error('The application version check could not be completed:', error)
        }
      } finally {
        checking = false
      }
    }

    void checkVersion()
    const intervalId = window.setInterval(() => {
      void checkVersion()
    }, 5 * 60 * 1000)

    return () => {
      active = false
      window.clearInterval(intervalId)
    }
  }, [])

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
            setSignedInFullName(name)
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
    if (!requestId || !organizationsLoaded || !contextLoaded) return

    let active = true

    const loadRequest = async () => {
      try {
        if (!requesterEmail) {
          setRequestAccessError('The signed-in email is unavailable, so access to this request could not be verified.')
          return
        }
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
            'RequesterEmail',
            'Approver',
            'ApproverEmail',
            'ApproverApproval',
            'Return',
            'Decline',
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
            setRequestAccessError(`The request could not be loaded: ${error?.message ?? 'The data source returned an error.'}`)
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
        const approversResult = await EFMModuleApproveService.getAll({
          select: ['GroupName', 'EmailToApprove'],
        })
        if (!approversResult.success) {
          throw new Error(`The approver list could not be loaded: ${approversResult.error?.message ?? 'The data source returned an error.'}`)
        }
        const groupResult = await GroupDetailsService.getAll({
          select: ['GroupName', 'MSGroupName', 'MSDistributionName'],
        })
        if (!groupResult.success) {
          throw new Error(`The approver group could not be loaded: ${groupResult.error?.message ?? 'The data source returned an error.'}`)
        }
        const approverGroups = groupResult.data.filter(
          (group) => group.MSDistributionName?.trim().toLowerCase() === request.ApproverEmail.trim().toLowerCase(),
        )
        const approverGroupNames = new Set(approverGroups.map((group) => group.GroupName.trim().toLowerCase()))
        const isRequester = request.RequesterEmail.trim().toLowerCase() === requesterEmail.trim().toLowerCase()
        const isAuthorizedApprover = approversResult.data.some(
          (approver) =>
            approverGroupNames.has(approver.GroupName.trim().toLowerCase()) &&
            approver.EmailToApprove.trim().toLowerCase() === requesterEmail.trim().toLowerCase(),
        )
        if (!isRequester && !isAuthorizedApprover) {
          setNotFoundRequestId(requestId)
          return
        }

        const organization = organizations.find(
          (item) => item.Id === request.MembershipOrganizationId,
        )
        setIsApprover(!isRequester && isAuthorizedApprover)
        setSavedRequest(request)
        setRequester(request.Requester)
        setSelectedOrganizationId(String(request.MembershipOrganizationId))
        setSponsor(organization?.InternalSponsor ?? '')
        setAmount(Number(request.Amount).toFixed(2))
        setStartDate(formatUsDate(request.StartDate))
        setEndDate(formatUsDate(request.EndDate))
        setMembershipDescription(organization?.Description ?? '')
      } catch (error) {
        if (active) {
          setRequestAccessError(`Access to this request could not be verified: ${describeError(error, 'The data source is unavailable.')}`)
        }
      } finally {
        if (active) setLoadedRequestId(requestId)
      }
    }

    void loadRequest()
    return () => {
      active = false
    }
  }, [contextLoaded, organizations, organizationsLoaded, requestId, requesterEmail])

  useEffect(() => {
    if (!savedRequest || !requestId) return

    let active = true
    void (async () => {
      try {
        const requestLogs = await getRequestWorkflowLogs(Number(requestId), await getRequestTableCodeId())
        if (!active) return
        setWorkflowLogs(requestLogs)
        setDescription(requestLogs[requestLogs.length - 1]?.NoteText ?? '')
      } catch (error) {
        if (active) {
          setPageError(`The workflow log could not be loaded: ${describeError(error, 'The data source is unavailable.')}`)
        }
      } finally {
        if (active) setLoadedWorkflowRequestId(requestId)
      }
    })()

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
    if (savedRequest) {
      if (!isApprover && !isReturnedRequester) return
      if (isApprover && approvalAction !== 'Process' && !description.trim()) {
        setNotesError('Notes is required')
        return
      }
      setNotesError('')
      if (!confirmed) {
        if (isReturnedRequester) {
          const errors = validate()
          setFieldErrors(errors)
          if (Object.keys(errors).length > 0) return
        }
        if (isReturnedRequester && !requesterEmail) {
          setPageError('Your signed-in email is unavailable. The request was not submitted.')
          return
        }
        setPageError('')
        setShowSubmitConfirmation(true)
        return
      }

      setShowSubmitConfirmation(false)
      setPageError('')
      setIsSubmitting(true)
      try {
        const requestKey = savedRequest.Id
        if (requestKey === undefined) {
          throw new Error('The request ID is unavailable, so the approval could not be saved.')
        }
        if (isReturnedRequester) {
          if (!requester || !requesterEmail) {
            throw new Error('Your signed-in name or email is unavailable. The request was not submitted.')
          }
          let resubmissionOrganizationId = Number(selectedOrganizationId)
          if (isNewOrganization) {
            const organizationResult = await MembershipOrganizationService.create({
              MemberOrganization: newOrganization.trim(),
              InternalSponsor: sponsor.trim(),
              Amount: Number(amount),
              Description: membershipDescription.trim() || undefined,
            })
            if (!organizationResult.success || organizationResult.data.Id === undefined) {
              throw new Error(`The new organization could not be saved: ${organizationResult.error?.message ?? 'The data source did not return an ID.'}`)
            }
            resubmissionOrganizationId = organizationResult.data.Id
          }
          const resubmissionTimestamp = getEasternDateTime(new Date())
          const updateResult = await PhilanthropyMembershipOrgRequestService.update(String(requestKey), {
            MembershipOrganizationId: resubmissionOrganizationId,
            Amount: Number(amount),
            StartDate: parseUsDate(startDate),
            EndDate: parseUsDate(endDate),
            Description: description.trim() || undefined,
            Return: false,
            UpdateAt: resubmissionTimestamp,
            UpdatedBy: requester,
          })
          if (!updateResult.success) {
            throw new Error(`The returned request could not be resubmitted: ${updateResult.error?.message ?? 'The data source returned an error.'}`)
          }
          const tableCodeResult = await TableCodeService.getAll({
            filter: "TableName eq 'PhilanthropyMembershipOrgRequest'",
            top: 1,
            select: ['TableCodeID', 'TableName'],
          })
          if (!tableCodeResult.success) {
            throw new Error(`The request was updated, but its workflow log could not be prepared: ${tableCodeResult.error?.message ?? 'The TableCode data source returned an error.'}`)
          }
          const tableCodeId = tableCodeResult.data.find(
            (tableCode) => tableCode.TableName === 'PhilanthropyMembershipOrgRequest',
          )?.TableCodeID
          if (tableCodeId === undefined) {
            throw new Error('The request was updated, but no matching TableCode was found for its workflow log.')
          }
          const workflowTimestamp = getEasternDateTime(new Date())
          const workflowResult = await WorkflowLogService.create({
            ActivityDate: workflowTimestamp,
            ActivityName: 'Request',
            ActivityUser: requester,
            ActivityAction: 'Resubmit',
            ProcInstID: requestKey,
            Status: 4,
            StartDate: workflowTimestamp,
            FinishDate: workflowTimestamp,
            TableCodeId: tableCodeId,
            NoteText: description.trim() || undefined,
            StatusDescription: `Waiting for ${savedRequest.Approver}`,
          })
          if (!workflowResult.success) {
            throw new Error(`The request was updated, but its workflow log could not be created: ${workflowResult.error?.message ?? 'The WorkflowLog data source returned an error.'}`)
          }
          const organization = isNewOrganization
            ? {
                MemberOrganization: newOrganization.trim(),
                InternalSponsor: sponsor.trim(),
                Description: membershipDescription.trim(),
              }
            : organizations.find((item) => item.Id === resubmissionOrganizationId)
          if (!organization) {
            throw new Error('The request was resubmitted, but its organization could not be loaded for the notification email.')
          }
          const [navigationResult, resubmissionLogs] = await Promise.all([
            NavigationService.getAll({ select: ['Title', 'NavURL'] }),
            getRequestWorkflowLogs(requestKey, tableCodeId),
          ])
          if (!navigationResult.success) {
            throw new Error(`The request was resubmitted, but its email link could not be loaded: ${navigationResult.error?.message ?? 'The Navigation data source returned an error.'}`)
          }
          const destination = navigationResult.data.find(
            (navigation) => navigation.Title === 'PhilanthropyMembership',
          )?.NavURL
          if (!destination) {
            throw new Error('The request was resubmitted, but no NavURL was found for PhilanthropyMembership.')
          }
          const url = new URL(destination, window.location.href)
          url.searchParams.set('requestid', String(requestKey))
          const emailBody = buildRequestEmailBody({
            href: url.href,
            organizationName: organization.MemberOrganization,
            requester,
            sponsor: organization.InternalSponsor,
            amount: Number(amount),
            startDate,
            endDate,
            description: organization.Description ?? '',
            nextDestination: savedRequest.Approver,
            requestId: String(requestKey),
            logs: resubmissionLogs,
          })
          const emailResults = await Promise.all([
            CaseyAdminFlowv1Service.Run({
              email: requesterEmail,
              email_1: '',
              text_1: 'Philanthropy Membership Org - Request',
              text: emailBody,
            }),
            CaseyAdminFlowv1Service.Run({
              email: savedRequest.ApproverEmail,
              email_1: '',
              text_1: 'Philanthropy Membership Org - Request',
              text: emailBody,
            }),
          ])
          const emailErrors = emailResults.flatMap((result, index) =>
            result.success
              ? []
              : [`${index === 0 ? 'requester' : 'approver'}: ${result.error?.message ?? 'Casey Admin Flow v1 returned an error.'}`],
          )
          if (emailErrors.length) {
            throw new Error(`The request was resubmitted, but email delivery failed for ${emailErrors.join('; ')}.`)
          }
          setSavedRequest({ ...savedRequest, MembershipOrganizationId: resubmissionOrganizationId, Amount: Number(amount), StartDate: parseUsDate(startDate), EndDate: parseUsDate(endDate), Description: description.trim() || undefined, Return: false })
          setWorkflowLogs((currentLogs) => [...currentLogs, workflowResult.data])
          return
        }
        const changes =
          approvalAction === 'Process'
            ? { ApproverApproval: true }
            : approvalAction === 'Decline'
              ? { Decline: true }
              : { Return: true }
        const updateResult = await PhilanthropyMembershipOrgRequestService.update(String(requestKey), changes)
        if (!updateResult.success) {
          throw new Error(`The approval action could not be saved: ${updateResult.error?.message ?? 'The data source returned an error.'}`)
        }
        setSavedRequest({ ...savedRequest, ...changes })

        const approvalTimestamp = getIndiaDateTime(new Date())
        let processLogError = ''
        let notificationLogs = workflowLogs
          .filter((log) => log.ProcInstID === requestKey)
          .sort((a, b) => (a.Id ?? 0) - (b.Id ?? 0))
        const tableCodeResult = await TableCodeService.getAll({
          filter: "TableName eq 'PhilanthropyMembershipOrgRequest'",
          top: 1,
          select: ['TableCodeID', 'TableName'],
        })
        const tableCodeId = tableCodeResult.success
          ? tableCodeResult.data.find(
              (tableCode) => tableCode.TableName === 'PhilanthropyMembershipOrgRequest',
            )?.TableCodeID
          : undefined
        if (!tableCodeResult.success) {
          processLogError = `The workflow log could not be prepared: ${tableCodeResult.error?.message ?? 'The TableCode data source returned an error.'}`
        } else if (tableCodeId === undefined) {
          processLogError = 'No matching TableCode was found for the workflow log.'
        }
        {
          if (!signedInFullName && !processLogError) {
            processLogError = 'Your signed-in name is unavailable, so the approval workflow log could not be created.'
          }
          if (!processLogError && tableCodeId !== undefined) {
            const workflowResult = await WorkflowLogService.create({
              ActivityDate: approvalTimestamp,
              ActivityName: 'Approval',
              ActivityUser: signedInFullName,
              ActivityAction: approvalAction,
              ProcInstID: requestKey,
              Status: 4,
              StartDate: approvalTimestamp,
              FinishDate: approvalTimestamp,
              TableCodeId: tableCodeId,
              NoteText: description.trim() || undefined,
              StatusDescription: approvalAction === 'Process'
                ? 'Process Complete'
                : approvalAction === 'Decline'
                  ? `Declined by ${signedInFullName}`
                  : `Returned by ${signedInFullName}`,
            })
            if (!workflowResult.success) {
              processLogError = `The workflow log could not be created: ${workflowResult.error?.message ?? 'The WorkflowLog data source returned an error.'}`
            } else {
              setWorkflowLogs((currentLogs) => [...currentLogs, workflowResult.data])
              if (approvalAction === 'Return') {
                setDescription(workflowResult.data.NoteText ?? '')
              }
            }
          }
        }
        if (tableCodeId !== undefined) {
          try {
            notificationLogs = await getRequestWorkflowLogs(requestKey, tableCodeId)
          } catch (error) {
            processLogError = `The workflow log records could not be loaded: ${describeError(error, 'The WorkflowLog data source is unavailable.')}`
          }
        }

        const organization = organizations.find(
          (item) => item.Id === savedRequest.MembershipOrganizationId,
        )
        if (!organization) {
          throw new Error('The request was updated, but its organization could not be loaded for the notification email.')
        }
        const navigationResult = await NavigationService.getAll({
          select: ['Title', 'NavURL'],
        })
        if (!navigationResult.success) {
          throw new Error(`The request was updated, but its email link could not be loaded: ${navigationResult.error?.message ?? 'The Navigation data source returned an error.'}`)
        }
        const destination = navigationResult.data.find(
          (navigation) => navigation.Title === 'PhilanthropyMembership',
        )?.NavURL
        if (!destination) {
          throw new Error('The request was updated, but no NavURL was found for PhilanthropyMembership.')
        }
        const url = new URL(destination, window.location.href)
        url.searchParams.set('requestid', String(requestKey))
        const emailBody = buildRequestEmailBody({
          href: url.href,
          organizationName: organization.MemberOrganization,
          requester: savedRequest.Requester,
          sponsor: organization.InternalSponsor ?? '',
          amount: Number(savedRequest.Amount),
          startDate: formatUsDate(savedRequest.StartDate),
          endDate: formatUsDate(savedRequest.EndDate),
          description: organization.Description ?? '',
          nextDestination: approvalAction === 'Return' ? savedRequest.Requester : undefined,
          requestId: String(requestKey),
          logs: notificationLogs.filter((log) => log.TableCodeId === tableCodeId),
        })
        const emailSubject = approvalAction === 'Decline' || approvalAction === 'Return'
          ? `Philanthropy Membership Org | ${approvalAction === 'Decline' ? 'Declined' : 'Returned'} by ${signedInFullName}`
          : 'Philanthropy Membership Org - Request'
        const emailResult = await CaseyAdminFlowv1Service.Run({
          email: savedRequest.RequesterEmail,
          email_1: '',
          text_1: emailSubject,
          text: emailBody,
        })
        if (!emailResult.success) {
          throw new Error(`The request was updated, but the requester email could not be sent: ${emailResult.error?.message ?? 'Casey Admin Flow v1 returned an error.'}`)
        }
        if (processLogError) {
          throw new Error(`The requester email was sent, but the approval workflow log failed: ${processLogError}`)
        }
      } catch (error) {
        setPageError(describeError(error, 'The approval action could not be completed.'))
      } finally {
        setIsSubmitting(false)
      }
      return
    }

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
          Description: membershipDescription.trim() || undefined,
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

      const distributionGroupName =
        !isNewOrganization &&
        Number(amount) === organizations.find((organization) => organization.Id === organizationId)?.Amount
          ? 'PhilanthropyGroup2'
          : 'PhilanthropyGroup1'
      const groupDetailsResult = await GroupDetailsService.getAll({
        filter: `MSGroupName eq '${distributionGroupName}'`,
        top: 1,
        select: ['GroupName', 'MSGroupName', 'MSDistributionName'],
      })
      if (!groupDetailsResult.success) {
        throw new Error(
          `The request could not be submitted because its approver and distribution email could not be loaded: ${groupDetailsResult.error?.message ?? 'The GroupDetails data source returned an error.'}`,
        )
      }
      const selectedGroup = groupDetailsResult.data.find(
        (group) => group.MSGroupName === distributionGroupName,
      )
      if (!selectedGroup?.GroupName) {
        throw new Error(`The request could not be submitted because no approver was found for ${distributionGroupName}.`)
      }
      const distributionEmail = selectedGroup.MSDistributionName
      if (!distributionEmail) {
        throw new Error(`The request could not be submitted because no distribution email was found for ${distributionGroupName}.`)
      }

      const requestTimestamp = getEasternDateTime(new Date())
      const requestResult = await PhilanthropyMembershipOrgRequestService.create({
        MembershipOrganizationId: organizationId,
        Requester: requester,
        Approver: selectedGroup.GroupName,
        ApproverEmail: distributionEmail,
        RequesterEmail: requesterEmail,
        RequestDate: requestTimestamp,
        Amount: Number(amount),
        Return: false,
        Decline: false,
        ApproverApproval: false,
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
      const organizationName = isNewOrganization
        ? newOrganization.trim()
        : organizations.find((organization) => organization.Id === organizationId)?.MemberOrganization
      if (!organizationName) {
        throw new Error('The request was saved, but the selected organization name could not be determined.')
      }
      const requestWorkflowLogs = await getRequestWorkflowLogs(requestResult.data.Id, tableCodeId)
      const emailBody = buildRequestEmailBody({
        href: url.href,
        organizationName,
        requester,
        sponsor: sponsor.trim(),
        amount: Number(amount),
        startDate,
        endDate,
        description: membershipDescription.trim(),
        nextDestination: selectedGroup.GroupName,
        requestId: createdRequestId,
        logs: requestWorkflowLogs,
      })
      const emailInput = {
        text_1: 'Philanthropy Membership Org - Request',
        text: emailBody,
        email_1: '',
      }
      const [requesterEmailResult, distributionEmailResult] = await Promise.all([
        CaseyAdminFlowv1Service.Run({ ...emailInput, email: requesterEmail }),
        CaseyAdminFlowv1Service.Run({ ...emailInput, email: distributionEmail }),
      ])
      const emailErrors = [
        requesterEmailResult.success
          ? null
          : `requester (${requesterEmail}): ${requesterEmailResult.error?.message ?? 'Casey Admin Flow v1 returned an error.'}`,
        distributionEmailResult.success
          ? null
          : `distribution group (${distributionEmail}): ${distributionEmailResult.error?.message ?? 'Casey Admin Flow v1 returned an error.'}`,
      ].filter((error): error is string => error !== null)
      if (emailErrors.length > 0) {
        throw new Error(`The request was saved, but email delivery failed for ${emailErrors.join('; ')}.`)
      }

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
    setHomeNavigationError('')
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
      const message = `The HR page could not be opened: ${describeError(error, 'The Navigation data source is unavailable.')}`
      setPageError(message)
      setHomeNavigationError(message)
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
      <main className="not-found-backdrop">
        <section className="form-panel not-found-panel" role="dialog" aria-modal="true" aria-labelledby="not-found-heading">
          <h1 id="not-found-heading">404 - Request Not Found</h1>
          <p>The requested form does not exist or has been deleted.</p>
          <a href="#" onClick={(event) => { event.preventDefault(); void handleClose() }}>Home</a>
          {homeNavigationError && <p className="page-error" role="alert">{homeNavigationError}</p>}
        </section>
      </main>
    )
  }

  if (requestAccessFailed) {
    return (
      <main className="not-found-backdrop">
        <section className="form-panel not-found-panel" role="alert">
          <h1>404 - Request Not Found</h1>
          <p>{requestAccessError}</p>
          <a href="#" onClick={(event) => { event.preventDefault(); void handleClose() }}>Home</a>
          {homeNavigationError && <p className="page-error" role="alert">{homeNavigationError}</p>}
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
        <h1>
          {isDeclined
            ? 'Philanthropy Membership - Declined'
            : isProcessed
            ? 'Philanthropy Membership - Processed'
            : savedRequest?.Return
              ? 'Philanthropy Membership - Return'
              : savedRequest
                ? 'Philanthropy Membership - Approval'
                : 'Philanthropy Membership - Request'}
        </h1>
        <nav className="header-nav" aria-label="Main navigation">
          <button type="button" onClick={() => setShowNavigationConfirmation(true)}>
            eForms
          </button>
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
                  disabled={Boolean(savedRequest) && !isReturnedRequester}
                  value={selectedOrganizationId}
                  onChange={(event) => {
                    const organizationId = event.target.value
                    const organization = organizations.find(
                      (item) => String(item.Id) === organizationId,
                    )
                    setSelectedOrganizationId(organizationId)
                    setSponsor(organization?.InternalSponsor ?? '')
                    setAmount(organization ? Number(organization.Amount).toFixed(2) : '')
                    setMembershipDescription(organization?.Description ?? '')
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
                    disabled={Boolean(savedRequest) && !isReturnedRequester}
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
                disabled={Boolean(savedRequest) && !isReturnedRequester}
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
                type="text"
                inputMode="decimal"
                disabled={Boolean(savedRequest) && !isReturnedRequester}
                value={amount}
                onChange={(event) => setAmount(event.target.value)}
                onBlur={() => {
                  if (amount.trim() && /^\d+(\.\d{1,2})?$/.test(amount)) {
                    setAmount(Number(amount).toFixed(2))
                  }
                }}
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
                  disabled={Boolean(savedRequest) && !isReturnedRequester}
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
                    disabled={Boolean(savedRequest) && !isReturnedRequester}
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
                  disabled={Boolean(savedRequest) && !isReturnedRequester}
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
                    disabled={Boolean(savedRequest) && !isReturnedRequester}
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

        <section className="form-panel" aria-labelledby="membership-description-heading">
          <h2 id="membership-description-heading">
            In a few sentences, please explain the value of this membership and to the work of others in the Foundation.
          </h2>
          <label className="visually-hidden" htmlFor="membership-description">
            Membership value and impact
          </label>
          <input
            id="membership-description"
            type="text"
            value={membershipDescription}
            disabled={!isNewOrganization}
            onChange={(event) => setMembershipDescription(event.target.value)}
          />
        </section>

        <section className="form-panel" aria-labelledby="description-heading">
          <h2 id="description-heading">Notes</h2>
          <label className="visually-hidden" htmlFor="description">Description</label>
          <input
            id="description"
            type="text"
            maxLength={250}
            disabled={isProcessed || isDeclined || (Boolean(savedRequest) && !isApprover && !isReturnedRequester)}
            value={description}
            onChange={(event) => {
              setDescription(event.target.value)
              if (event.target.value.trim()) setNotesError('')
            }}
          />
          {notesError && <span className="field-error notes-error" role="alert">{notesError}</span>}
        </section>

        <section className="form-panel form-actions">
          <p>Please click <strong>Submit</strong> when you are ready to submit this form.</p>
          {pageError && <p className="page-error" role="alert">{pageError}</p>}
          <div className="action-row">
            {savedRequest && isApprover && !isProcessed && !isDeclined && !savedRequest.Return && (
              <div className="approval-action">
                <label className="field-label" htmlFor="approval-action">Action</label>
                <select
                  id="approval-action"
                  value={approvalAction}
                  onChange={(event) => setApprovalAction(event.target.value)}
                  disabled={isSubmitting}
                >
                  <option value="Process">Process</option>
                  <option value="Decline">Decline</option>
                  <option value="Return">Return</option>
                </select>
              </div>
            )}
            <div className="button-row">
              {(!savedRequest || (isApprover && !isProcessed && !isDeclined && !savedRequest.Return) || (isReturnedRequester && !isDeclined)) && (
                <button
                  type="submit"
                  disabled={isSubmitting || (!savedRequest && (!requester || !requesterEmail))}
                >
                  Submit
                </button>
              )}
              <button
                className="secondary-button"
                type="button"
                onClick={handleClose}
                disabled={isClosing}
              >
                {isClosing ? 'Closing…' : 'Close'}
              </button>
            </div>
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
                      hourCycle: 'h23',
                    }).replace(',', '')}</td>
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
              {savedRequest
                ? `Please confirm the ${approvalAction.toLowerCase()} action for this request.`
                : 'Please confirm if you are ready to submit the eForm.'}
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
      {showNavigationConfirmation && (
        <div className="confirmation-backdrop">
          <section
            className="confirmation-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="navigation-confirmation-message"
          >
            <p id="navigation-confirmation-message">
              Are you sure you want to leave this form?
            </p>
            <div className="confirmation-actions">
              <button
                type="button"
                onClick={() => {
                  setShowNavigationConfirmation(false)
                  void handleClose()
                }}
              >
                Ok
              </button>
              <button
                className="cancel-button"
                type="button"
                onClick={() => setShowNavigationConfirmation(false)}
              >
                Cancel
              </button>
            </div>
          </section>
        </div>
      )}
      <footer className="app-version">
        App version: {databaseVersion ?? __APP_VERSION__}
      </footer>
    </main>
  )
}

export default App
