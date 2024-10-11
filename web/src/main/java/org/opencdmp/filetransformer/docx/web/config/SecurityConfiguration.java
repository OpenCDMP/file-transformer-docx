package org.opencdmp.filetransformer.docx.web.config;

import gr.cite.commons.web.oidc.configuration.WebSecurityProperties;
import gr.cite.commons.web.oidc.configuration.filter.ApiKeyFilter;
import jakarta.servlet.http.HttpServletRequest;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.authentication.AuthenticationManagerResolver;
import org.springframework.security.config.Customizer;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.config.annotation.web.configurers.AbstractHttpConfigurer;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.authentication.preauth.AbstractPreAuthenticatedProcessingFilter;

import java.util.Set;

@Configuration
@EnableWebSecurity
public class SecurityConfiguration {

    private final ApiKeyFilter apiKeyFilter;
    private final WebSecurityProperties webSecurityProperties;
    private final AuthenticationManagerResolver<HttpServletRequest> authenticationManagerResolver;

    @Autowired
    public SecurityConfiguration(ApiKeyFilter apiKeyFilter, WebSecurityProperties webSecurityProperties, AuthenticationManagerResolver<HttpServletRequest> authenticationManagerResolver) {
        this.apiKeyFilter = apiKeyFilter;
        this.webSecurityProperties = webSecurityProperties;
        this.authenticationManagerResolver = authenticationManagerResolver;
    }

    @Bean
    protected SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
        if (webSecurityProperties.isEnabled()) {
            http.csrf(AbstractHttpConfigurer::disable)
                    .cors(Customizer.withDefaults())
                    .addFilterBefore(apiKeyFilter, AbstractPreAuthenticatedProcessingFilter.class)
                    .authorizeHttpRequests(authorizationManagerRequestMatcherRegistry -> authorizationManagerRequestMatcherRegistry
                            .requestMatchers(buildAntPatterns(webSecurityProperties.getAuthorizedEndpoints())).authenticated()
                            .requestMatchers(buildAntPatterns(webSecurityProperties.getAllowedEndpoints())).anonymous())
                    .sessionManagement(httpSecuritySessionManagementConfigurer -> httpSecuritySessionManagementConfigurer.sessionCreationPolicy(SessionCreationPolicy.NEVER))
                    .oauth2ResourceServer(oauth2 -> oauth2.authenticationManagerResolver(authenticationManagerResolver));
            return http.build();
        } else {
            return http.csrf(AbstractHttpConfigurer::disable)
                    .cors(Customizer.withDefaults())
                    .authorizeHttpRequests(authorizationManagerRequestMatcherRegistry ->
                            authorizationManagerRequestMatcherRegistry.anyRequest().anonymous())
                    .build();
        }
    }

    private String[] buildAntPatterns(Set<String> endpoints) {
        if (endpoints == null) {
            return new String[0];
        }
        return endpoints.stream()
                .filter(endpoint -> endpoint != null && !endpoint.isBlank())
                .map(endpoint -> "/" + stripUnnecessaryCharacters(endpoint) + "/**")
                .toArray(String[]::new);
    }

    private String stripUnnecessaryCharacters(String endpoint) {
        endpoint = endpoint.strip();
        if (endpoint.startsWith("/")) {
            endpoint = endpoint.substring(1);
        }
        if (endpoint.endsWith("/")) {
            endpoint = endpoint.substring(0, endpoint.length() - 1);
        }
        return endpoint;
    }
}
