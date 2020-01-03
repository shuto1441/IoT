#!/usr/bin/env python
# -*- coding: utf-8 -*-
import slackweb
slack = slackweb.Slack(url="https://hooks.slack.com/services/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")
slack.notify(text="せいや")
"""
attachments = []
attachment = {"title": "大会進行状況",
                "pretext": "現在の大会進行状況をお知らせします",
                "text": "あ",
                "mrkdwn_in": ["text", "pretext"]}
attachments.append(attachment)
slack.notify(attachments=attachments)
"""