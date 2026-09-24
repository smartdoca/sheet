import { Chart, type G2Spec } from '@antv/g2'
import { useEffect, useRef } from 'react'

import type { AnalysisChartData } from './types'

export const ANALYSIS_CHART_COMPONENT = 'uos.analysis-chart'
export const ANALYSIS_CHART_REMOVE_EVENT = 'uos:remove-analysis-chart'

function buildOptions(data: AnalysisChartData) {
  const common = {
    data: data.values,
    animate: false,
    legend: { color: { position: 'bottom' } },
    scale: { color: { palette: 'category10' } },
  }

  if (data.type === 'pie' || data.type === 'donut') {
    const firstSeries = data.values[0]?.series
    return {
      type: 'interval',
      data: data.values.filter((item) => item.series === firstSeries),
      coordinate: {
        type: 'theta',
        innerRadius: data.type === 'donut' ? 0.55 : 0,
      },
      transform: [{ type: 'stackY' }],
      encode: { y: 'value', color: 'category' },
      legend: { color: { position: 'bottom' } },
      labels: [{ text: 'category', position: 'outside' }],
      animate: false,
    }
  }

  if (data.type === 'scatter') {
    return {
      type: 'point',
      encode: { x: 'x', y: 'y', color: 'series', shape: 'point' },
      style: { r: 5 },
      ...common,
    }
  }

  if (data.type === 'radar') {
    return {
      type: 'line',
      ...common,
      coordinate: { type: 'polar' },
      encode: { x: 'category', y: 'value', color: 'series' },
      style: { lineWidth: 2 },
    }
  }

  if (data.type === 'area') {
    return {
      type: 'view',
      ...common,
      children: [
        {
          type: 'area',
          encode: { x: 'category', y: 'value', color: 'series' },
          style: { fillOpacity: 0.3 },
        },
        {
          type: 'line',
          encode: { x: 'category', y: 'value', color: 'series' },
        },
      ],
    }
  }

  if (data.type === 'line') {
    return {
      type: 'line',
      ...common,
      encode: { x: 'category', y: 'value', color: 'series' },
      style: { lineWidth: 2 },
    }
  }

  return {
    type: 'interval',
    ...common,
    coordinate: data.type === 'bar' ? { transform: [{ type: 'transpose' }] } : undefined,
    encode: { x: 'category', y: 'value', color: 'series' },
    transform: [{ type: 'dodgeX' }],
  }
}

export function AnalysisChart({ data: rawData }: { data?: unknown }) {
  const containerRef = useRef<HTMLDivElement>(null)
  const data = rawData as AnalysisChartData | undefined

  useEffect(() => {
    if (!containerRef.current || !data) return
    const chart = new Chart({ container: containerRef.current, autoFit: true })
    chart.options(buildOptions(data) as G2Spec)
    void chart.render().then(() => chart.forceFit())
    const observer = new ResizeObserver(([entry]) => {
      const { width, height } = entry.contentRect
      if (width > 0 && height > 0) void chart.changeSize(width, height)
    })
    observer.observe(containerRef.current)
    return () => {
      observer.disconnect()
      chart.destroy()
    }
  }, [data])

  if (!data) return null

  return (
    <article className="uos-analysis-chart" data-chart-id={data.chartId}>
      <header>
        <div>
          <strong>{data.title}</strong>
          <span>{data.sourceRange}</span>
        </div>
        <button
          type="button"
          title={data.removeLabel ?? 'Remove chart'}
          aria-label={data.removeLabel ?? 'Remove chart'}
          onClick={() => window.dispatchEvent(new CustomEvent(ANALYSIS_CHART_REMOVE_EVENT, {
            detail: { chartId: data.chartId, workbookId: data.workbookId },
          }))}
        >
          ×
        </button>
      </header>
      <div ref={containerRef} className="uos-analysis-chart__canvas" />
    </article>
  )
}
