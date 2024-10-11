package org.opencdmp.filetransformer.docx.service.pdf;

import gr.cite.tools.logging.LoggerService;
import gr.cite.tools.logging.MapLogEntry;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.MediaType;
import org.springframework.http.client.MultipartBodyBuilder;
import org.springframework.stereotype.Component;
import org.springframework.web.reactive.function.BodyInserters;
import org.springframework.web.reactive.function.client.ExchangeFilterFunction;
import org.springframework.web.reactive.function.client.WebClient;
import reactor.core.publisher.Mono;

import java.util.UUID;

@Component
public class PdfServiceImpl implements PdfService {
	private static final LoggerService logger = new LoggerService(LoggerFactory.getLogger(PdfServiceImpl.class));

    private final PdfServiceProperties pdfServiceProperties;

	public PdfServiceImpl(PdfServiceProperties pdfServiceProperties) {
		this.pdfServiceProperties = pdfServiceProperties;
	}

	@Override
    public byte[] convertToPDF(byte[] file) {
        WebClient webClient = WebClient.builder().filters(exchangeFilterFunctions -> {
	        exchangeFilterFunctions.add(logRequest());
	        exchangeFilterFunctions.add(logResponse());
        }).baseUrl(pdfServiceProperties.getUrl()) .codecs(codecs -> codecs
	        .defaultCodecs()
	        .maxInMemorySize(this.pdfServiceProperties.getMaxInMemorySizeInBytes())
        ).build();
		MultipartBodyBuilder builder = new MultipartBodyBuilder();
		builder.part("files", new ByteArrayResource(file)).filename(UUID.randomUUID() + ".docx");

		return webClient.post().uri("forms/libreoffice/convert")
                .headers(httpHeaders -> {
	                httpHeaders.setContentType(MediaType.MULTIPART_FORM_DATA);
	                httpHeaders.add("Content-disposition", "attachment; filename=" + UUID.randomUUID() + ".pdf");
	                httpHeaders.add("Content-type", "application/pdf");
                })
                .body(BodyInserters.fromMultipartData(builder.build()))
                .retrieve().bodyToMono(byte[].class).block();
    }



	private static ExchangeFilterFunction logRequest() {
		return ExchangeFilterFunction.ofRequestProcessor(clientRequest -> {
			logger.debug(new MapLogEntry("Request").And("method", clientRequest.method().toString()).And("url", clientRequest.url()));
			return Mono.just(clientRequest);
		});
	}

	private static ExchangeFilterFunction logResponse() {
		return ExchangeFilterFunction.ofResponseProcessor(response -> {
			if (response.statusCode().isError()) {
				return response.mutate().build().bodyToMono(String.class)
						.flatMap(body -> {
							logger.error(new MapLogEntry("Response").And("method", response.request().getMethod().toString()).And("url", response.request().getURI()).And("status", response.statusCode()).And("body", body));
							return Mono.just(response);
						});
			}
			return Mono.just(response);

		});
	}
}
