// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.Agents.Builder.App;
using Microsoft.Extensions.Logging;
using Microsoft.Teams.Apps;
using System.Linq;

namespace TeamsExtension;

/// <summary>
/// AgentApplication helper for registering <see cref="TeamsSdkAgentExtension"/>.
/// </summary>
public static class AgentApplicationTeamsExtension
{
    /// <summary>
    /// Registers the Teams SDK routing extension on the provided <see cref="AgentApplication"/>.
    /// </summary>
    public static AgentApplication UseTeamsExtension(
        this AgentApplication agentApplication,
        TeamsBotApplication teamsBotApplication,
        ILogger<TeamsSdkAgentExtension>? logger = null)
    {
        if (agentApplication.RegisteredExtensions.OfType<TeamsSdkAgentExtension>().Any())
        {
            return agentApplication;
        }

        var extension = new TeamsSdkAgentExtension(agentApplication, teamsBotApplication, logger);
        agentApplication.RegisterExtension(extension, _ => { });
        return agentApplication;
    }
}
