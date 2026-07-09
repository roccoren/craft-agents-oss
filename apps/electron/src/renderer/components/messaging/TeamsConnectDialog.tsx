/**
 * TeamsConnectDialog — Azure Bot (Bot Framework) connect flow for Microsoft
 * Teams. Collects App ID + App Password (+ optional Tenant ID) and a public
 * messaging endpoint (bring-your-own HTTPS URL). After save, shows the
 * `<publicUrl>/api/messages` value to paste into the Azure Bot resource.
 */
import * as React from 'react'
import { Check, X, Copy } from 'lucide-react'
import { useTranslation } from 'react-i18next'
import { toast } from 'sonner'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogFooter,
} from '@/components/ui/dialog'
import { Button } from '@/components/ui/button'
import { Spinner } from '@craft-agent/ui'
import { SettingsSecretInput } from '@/components/settings'

interface TeamsConnectDialogProps {
  open: boolean
  onOpenChange: (open: boolean) => void
  reconfigure?: boolean
  onSaved?: () => void
}

type TestResult =
  | { state: 'idle' }
  | { state: 'testing' }
  | { state: 'success' }
  | { state: 'error'; error: string }

export function TeamsConnectDialog({ open, onOpenChange, reconfigure = false, onSaved }: TeamsConnectDialogProps) {
  const { t } = useTranslation()
  const [appId, setAppId] = React.useState('')
  const [appPassword, setAppPassword] = React.useState('')
  const [tenantId, setTenantId] = React.useState('')
  const [byoUrl, setByoUrl] = React.useState('')
  const [saving, setSaving] = React.useState(false)
  const [test, setTest] = React.useState<TestResult>({ state: 'idle' })
  const [endpoint, setEndpoint] = React.useState<string | null>(null)

  React.useEffect(() => {
    if (!open) {
      setAppId('')
      setAppPassword('')
      setTenantId('')
      setByoUrl('')
      setTest({ state: 'idle' })
      setSaving(false)
      setEndpoint(null)
    }
  }, [open])

  const ready = appId.trim().length > 0 && appPassword.trim().length > 0
  const canSave = ready && byoUrl.trim().length > 0 && test.state === 'success'

  const handleTest = async () => {
    if (!ready) return
    setTest({ state: 'testing' })
    try {
      const result = await window.electronAPI.testTeamsCredentials({
        appId: appId.trim(),
        appPassword: appPassword.trim(),
        tenantId: tenantId.trim() || undefined,
      })
      setTest(result.success ? { state: 'success' } : { state: 'error', error: result.error ?? t('common.error') })
    } catch (err) {
      setTest({ state: 'error', error: err instanceof Error ? err.message : t('common.error') })
    }
  }

  const handleSave = async () => {
    if (!canSave) return
    setSaving(true)
    try {
      const res = await window.electronAPI.saveTeamsCredentials({
        appId: appId.trim(),
        appPassword: appPassword.trim(),
        tenantId: tenantId.trim() || undefined,
        tunnelMode: 'byo',
        byoUrl: byoUrl.trim(),
      })
      setEndpoint(res.messagingEndpoint)
      toast.success(t('settings.messaging.teams.saved'))
      onSaved?.()
    } catch (err) {
      toast.error(err instanceof Error ? err.message : t('settings.messaging.teams.saveFailed'))
    } finally {
      setSaving(false)
    }
  }

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-[520px]">
        <DialogHeader>
          <DialogTitle>
            {reconfigure
              ? t('settings.messaging.teams.reconfigureTitle')
              : t('settings.messaging.teams.connectTitle')}
          </DialogTitle>
          <DialogDescription className="whitespace-pre-line">
            {t('settings.messaging.teams.instructions')}
          </DialogDescription>
        </DialogHeader>

        <div className="space-y-3 py-2">
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.appIdLabel')}</div>
            <SettingsSecretInput
              value={appId}
              onChange={setAppId}
              placeholder={t('settings.messaging.teams.appIdPlaceholder')}
              disabled={saving}
            />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.appPasswordLabel')}</div>
            <SettingsSecretInput
              value={appPassword}
              onChange={setAppPassword}
              placeholder={t('settings.messaging.teams.appPasswordPlaceholder')}
              disabled={saving}
            />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.tenantIdLabel')}</div>
            <SettingsSecretInput
              value={tenantId}
              onChange={setTenantId}
              placeholder={t('settings.messaging.teams.tenantIdPlaceholder')}
              disabled={saving}
            />
          </div>
          <div>
            <div className="mb-1.5 text-xs text-muted-foreground">{t('settings.messaging.teams.byoUrlLabel')}</div>
            <SettingsSecretInput
              value={byoUrl}
              onChange={setByoUrl}
              placeholder={t('settings.messaging.teams.byoUrlPlaceholder')}
              disabled={saving}
            />
          </div>

          <div className="flex items-center gap-2">
            <Button
              variant="outline"
              size="sm"
              onClick={handleTest}
              disabled={!ready || test.state === 'testing' || saving}
            >
              {test.state === 'testing' && <Spinner className="mr-1 text-[14px]" />}
              {t('settings.messaging.teams.testConnection')}
            </Button>
            {test.state === 'success' && (
              <span className="inline-flex items-center gap-1 text-xs text-emerald-600 dark:text-emerald-400">
                <Check className="h-3.5 w-3.5" />
                {t('settings.messaging.teams.testOk')}
              </span>
            )}
            {test.state === 'error' && (
              <span className="inline-flex items-center gap-1 text-xs text-destructive">
                <X className="h-3.5 w-3.5" />
                {test.error}
              </span>
            )}
          </div>

          {endpoint && (
            <div className="rounded-md border border-border bg-muted/40 p-2 text-xs">
              <div className="mb-1 font-medium">{t('settings.messaging.teams.endpointTitle')}</div>
              <div className="flex items-center gap-2">
                <code className="truncate">{endpoint}</code>
                <Button
                  variant="ghost"
                  size="icon"
                  className="h-6 w-6"
                  onClick={() => {
                    void navigator.clipboard.writeText(endpoint)
                    toast.success(t('common.copied'))
                  }}
                >
                  <Copy className="h-3.5 w-3.5" />
                </Button>
              </div>
              <div className="mt-1 text-muted-foreground">{t('settings.messaging.teams.endpointHint')}</div>
            </div>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={() => onOpenChange(false)} disabled={saving}>
            {endpoint ? t('common.done') : t('common.cancel')}
          </Button>
          {!endpoint && (
            <Button variant="outline" size="sm" onClick={handleSave} disabled={!canSave || saving}>
              {saving && <Spinner className="mr-1 text-[14px]" />}
              {t('settings.messaging.teams.save')}
            </Button>
          )}
        </DialogFooter>
      </DialogContent>
    </Dialog>
  )
}
